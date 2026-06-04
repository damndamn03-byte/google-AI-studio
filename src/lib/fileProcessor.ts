import JSZip from 'jszip';
import * as pdfjsLib from 'pdfjs-dist';

// 使用 unpkg 載入對應版本的 worker，這在大多數環境下比 cdnjs 更穩定
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;

export interface ProcessingResult {
  fileName: string;
  filePath: string;
  status: 'success' | 'skipped_scan' | 'error' | 'legacy_skipped';
  imagesExtracted: number;
  message?: string;
}

export async function extractImagesFromOffice(file: File, targetDirHandle: FileSystemDirectoryHandle): Promise<number> {
  const arrayBuffer = await file.arrayBuffer();
  
  // 確保在任何打包或載入環境中都能正確獲得 JSZip 的構造函數
  const JSZipConstructor = (JSZip as any).default || JSZip;
  if (!JSZipConstructor || typeof JSZipConstructor.loadAsync !== 'function') {
    throw new Error('JSZip 載入失敗或損毀，無法解析 Office 檔案格式。');
  }
  
  const zip = await JSZipConstructor.loadAsync(arrayBuffer);
  
  const nameLower = file.name.toLowerCase();
  const isDocx = nameLower.endsWith('.docx');
  const isXlsx = nameLower.endsWith('.xlsx');
  
  const parser = new DOMParser();
  const orderedMediaPaths: string[] = [];
  
  if (isXlsx) {
    try {
      // 1. 取得所有 Worksheet 檔案並按數字進行自然排序 (e.g. sheet1.xml, sheet2.xml)
      const sheetFiles = Object.keys(zip.files).filter(path => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path));
      sheetFiles.sort((a, b) => {
        const matchA = a.match(/\d+/);
        const matchB = b.match(/\d+/);
        const numA = matchA ? parseInt(matchA[0], 10) : 0;
        const numB = matchB ? parseInt(matchB[0], 10) : 0;
        return numA - numB;
      });
      
      // 2. 依照工作表順序，查找關聯的 Drawing 檔案 (sheet rels)
      const drawingOrder: string[] = [];
      for (const sheetPath of sheetFiles) {
        const sheetName = sheetPath.split('/').pop()!;
        const relsPath = `xl/worksheets/_rels/${sheetName}.rels`;
        if (zip.file(relsPath)) {
          const relsText = await zip.file(relsPath)!.async('string');
          const relsDoc = parser.parseFromString(relsText, 'application/xml');
          const relationships = relsDoc.getElementsByTagName('Relationship');
          for (let i = 0; i < relationships.length; i++) {
            const rel = relationships[i];
            const type = rel.getAttribute('Type');
            const target = rel.getAttribute('Target');
            if (type?.endsWith('/drawing') && target) {
              let resolvedTarget = target;
              if (target.startsWith('../')) {
                resolvedTarget = 'xl/' + target.substring(3);
              } else {
                resolvedTarget = 'xl/worksheets/' + target;
              }
              if (!drawingOrder.includes(resolvedTarget)) {
                drawingOrder.push(resolvedTarget);
              }
            }
          }
        }
      }
      
      // 3. 備用方案：將其餘 XML drawings 按數值自然排序，追加到最後
      const allDrawings = Object.keys(zip.files).filter(path => /^xl\/drawings\/drawing\d+\.xml$/i.test(path));
      allDrawings.sort((a, b) => {
        const matchA = a.match(/\d+/);
        const matchB = b.match(/\d+/);
        const numA = matchA ? parseInt(matchA[0], 10) : 0;
        const numB = matchB ? parseInt(matchB[0], 10) : 0;
        return numA - numB;
      });
      for (const dr of allDrawings) {
        if (!drawingOrder.includes(dr)) {
          drawingOrder.push(dr);
        }
      }
      
      interface ExcelImageInfo {
        drawingIndex: number;
        row: number;
        col: number;
        mediaPath: string;
      }
      
      const collectedImages: ExcelImageInfo[] = [];
      
      const findInSubtree = (parent: Element, name: string): Element | null => {
        const elements = parent.getElementsByTagName('*');
        for (let i = 0; i < elements.length; i++) {
          if (elements[i].localName === name) {
            return elements[i];
          }
        }
        return null;
      };
      
      // 4. 解析各繪圖 XML 當中的儲存格位置 (Row 與 Col)
      for (let d = 0; d < drawingOrder.length; d++) {
        const drawingPath = drawingOrder[d];
        const drawingXmlFile = zip.file(drawingPath);
        if (!drawingXmlFile) continue;
        
        const drawingText = await drawingXmlFile.async('string');
        const lastSlash = drawingPath.lastIndexOf('/');
        const dir = drawingPath.substring(0, lastSlash);
        const fileName = drawingPath.substring(lastSlash + 1);
        const relsPath = `${dir}/_rels/${fileName}.rels`;
        
        const rIdToMedia = new Map<string, string>();
        if (zip.file(relsPath)) {
          const relsText = await zip.file(relsPath)!.async('string');
          const relsDoc = parser.parseFromString(relsText, 'application/xml');
          const relationships = relsDoc.getElementsByTagName('Relationship');
          for (let i = 0; i < relationships.length; i++) {
            const rel = relationships[i];
            const id = rel.getAttribute('Id');
            let target = rel.getAttribute('Target');
            if (id && target) {
              if (target.startsWith('../')) {
                target = 'xl/' + target.substring(3); // ../media/image1.png -> xl/media/image1.png
              } else if (!target.startsWith('xl/')) {
                target = 'xl/drawings/' + target;
              }
              rIdToMedia.set(id, target);
            }
          }
        }
        
        const doc = parser.parseFromString(drawingText, 'application/xml');
        const anchors = Array.from(doc.getElementsByTagName('*')).filter(el => el.localName?.endsWith('Anchor'));
        
        for (const anchor of anchors) {
          const fromNode = findInSubtree(anchor, 'from');
          let row = 0;
          let col = 0;
          if (fromNode) {
            const rowNode = findInSubtree(fromNode, 'row');
            const colNode = findInSubtree(fromNode, 'col');
            row = rowNode ? parseInt(rowNode.textContent || '0', 10) : 0;
            col = colNode ? parseInt(colNode.textContent || '0', 10) : 0;
          }
          
          const blipNode = findInSubtree(anchor, 'blip');
          let rId = null;
          if (blipNode) {
            rId = blipNode.getAttribute('r:embed') || blipNode.getAttribute('embed');
            if (!rId) {
              // 備份命名空間查詢
              rId = blipNode.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
            }
            if (!rId) {
              for (let i = 0; i < blipNode.attributes.length; i++) {
                if (blipNode.attributes[i].localName === 'embed') {
                  rId = blipNode.attributes[i].value;
                  break;
                }
              }
            }
          }
          
          if (rId) {
            const mediaPath = rIdToMedia.get(rId);
            if (mediaPath && zip.file(mediaPath)) {
              collectedImages.push({
                drawingIndex: d,
                row,
                col,
                mediaPath
              });
            }
          }
        }
      }
      
      // 5. 按工作表排序 -> 按橫列Row優先 (一列擷取完再換下一列) -> 再按直欄Col 排序依序呈現
      collectedImages.sort((a, b) => {
        if (a.drawingIndex !== b.drawingIndex) {
          return a.drawingIndex - b.drawingIndex;
        }
        if (a.row !== b.row) {
          return a.row - b.row;
        }
        return a.col - b.col;
      });
      
      // 將排序後的媒體路徑加入 (不進行去重，以支援重複位置的同一個圖片能依序、個別擷取)
      for (const img of collectedImages) {
        orderedMediaPaths.push(img.mediaPath);
      }
    } catch (e) {
      console.warn('XLSX 圖片儲存格排序解析失敗，將退回傳統處理方式', e);
    }
  } else if (isDocx) {
    try {
      // 1. 載入 Word 關聯對應
      const rIdToMedia = new Map<string, string>();
      const relsPath = 'word/_rels/document.xml.rels';
      if (zip.file(relsPath)) {
        const relsText = await zip.file(relsPath)!.async('string');
        const relsDoc = parser.parseFromString(relsText, 'application/xml');
        const relationships = relsDoc.getElementsByTagName('Relationship');
        for (let i = 0; i < relationships.length; i++) {
          const rel = relationships[i];
          const id = rel.getAttribute('Id');
          let target = rel.getAttribute('Target');
          if (id && target) {
            if (!target.startsWith('word/')) {
              target = 'word/' + target;
            }
            rIdToMedia.set(id, target);
          }
        }
      }
      
      // 2. 載入並解析主文檔 XML (word/document.xml) 內 blip 出現之物理順序
      const docPath = 'word/document.xml';
      if (zip.file(docPath)) {
        const docText = await zip.file(docPath)!.async('string');
        const doc = parser.parseFromString(docText, 'application/xml');
        const blips = doc.getElementsByTagNameNS('*', 'blip');
        
        for (let i = 0; i < blips.length; i++) {
          const blip = blips[i];
          let rId = blip.getAttribute('r:embed') || blip.getAttribute('embed');
          if (!rId) {
            for (let a = 0; a < blip.attributes.length; a++) {
              if (blip.attributes[a].localName === 'embed') {
                rId = blip.attributes[a].value;
                break;
              }
            }
          }
          if (rId) {
            const mediaPath = rIdToMedia.get(rId);
            if (mediaPath && zip.file(mediaPath)) {
              orderedMediaPaths.push(mediaPath);
            }
          }
        }
        
        // 額外使用正則表達式作為強固安全保障 (僅在 DOM 解析無效時，才當作備用覆蓋層，避免多重探針產生衝突)
        if (orderedMediaPaths.length === 0) {
          const matches = [...docText.matchAll(/(?:r:embed|embed)=["']([^"']+)["']/g)];
          for (const m of matches) {
            const rId = m[1];
            const mediaPath = rIdToMedia.get(rId);
            if (mediaPath && zip.file(mediaPath)) {
              orderedMediaPaths.push(mediaPath);
            }
          }
        }
      }
    } catch (e) {
      console.warn('DOCX 圖片文檔流順序解析失敗，將退回傳統處理方式', e);
    }
  }
  
  // 6. 計算本 Office 的媒體資料夾基礎路徑
  let mediaPathPrefix = 'xl/media/';
  if (isDocx) {
    mediaPathPrefix = 'word/media/';
  } else {
    // 自動檢測媒體資料夾以相容多種變體
    const firstMediaFolder = Object.keys(zip.files).find(p => p.includes('/media/'));
    if (firstMediaFolder) {
      const idx = firstMediaFolder.indexOf('/media/');
      mediaPathPrefix = firstMediaFolder.substring(0, idx + 7);
    }
  }
  
  // 7. 彙整 ZIP 中所有媒體檔案，確保 100% 都不缺失 (可能存在未被 document/sheet XML 列出的圖片，例如：頁首頁尾、浮水印、背景)
  const allMediaFilesInZip = Object.keys(zip.files).filter(path => path.startsWith(mediaPathPrefix) && !zip.files[path].dir);
  
  const finalPathsToExtract: string[] = [...orderedMediaPaths];
  for (const path of allMediaFilesInZip) {
    if (!finalPathsToExtract.includes(path)) {
      finalPathsToExtract.push(path);
    }
  }
  
  let count = 0;
  for (let i = 0; i < finalPathsToExtract.length; i++) {
    const path = finalPathsToExtract[i];
    try {
      const zipFile = zip.files[path];
      if (!zipFile || zipFile.dir) continue;
      
      const blob = await zipFile.async('blob');
      const originalName = path.split('/').pop() || `image_${count + 1}`;
      
      // 圖片命名規則：前置三位數補零的序列（例如：001_image1.png），保證視窗與電腦檔案總管能依序排列
      const prefix = String(count + 1).padStart(3, '0');
      const fileName = `${prefix}_${originalName}`;
      
      const fileHandle = await targetDirHandle.getFileHandle(fileName, { create: true });
      const writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();
      count++;
    } catch (e) {
      console.warn(`跳過特定圖片擷取，路徑: ${path}`, e);
    }
  }
  
  return count;
}

export async function extractImagesFromPdf(file: File, targetDirHandle: FileSystemDirectoryHandle): Promise<{ count: number; isScanned: boolean }> {
  const arrayBuffer = await file.arrayBuffer();
  
  const loadingTask = pdfjsLib.getDocument({ 
    data: arrayBuffer,
  });
  
  const pdf = await loadingTask.promise;
  
  let totalImages = 0;
  let totalTextChars = 0;
  const pagesToCheck = Math.min(pdf.numPages, 3);
  
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    
    if (i <= pagesToCheck) {
      const textContent = await page.getTextContent();
      const text = textContent.items.map((item: any) => item.str).join('');
      totalTextChars += text.trim().length;
    }
    
    const operatorList = await page.getOperatorList();
    const validImages = [
      pdfjsLib.OPS.paintImageXObject,
      pdfjsLib.OPS.paintInlineImageXObject,
      pdfjsLib.OPS.paintImageXObjectRepeat
    ];

    for (let j = 0; j < operatorList.fnArray.length; j++) {
      if (validImages.includes(operatorList.fnArray[j])) {
        const objId = operatorList.argsArray[j][0];
        
        try {
          // 嘗試從 page.objs 或 page.commonObjs 獲取圖片對象
          let img: any;
          try {
            img = await page.objs.get(objId);
          } catch (e) {
            img = await page.commonObjs.get(objId);
          }

          if (img && (img.data || img.bitmap)) {
            const blob = await imageToBlob(img);
            if (blob) {
              const fileName = `page${i}_image_${totalImages + 1}.png`;
              const fileHandle = await targetDirHandle.getFileHandle(fileName, { create: true });
              const writable = await fileHandle.createWritable();
              await writable.write(blob);
              await writable.close();
              totalImages++;
            }
          }
        } catch (e) {
          console.warn(`第 ${i} 頁圖片擷取失敗:`, e);
        }
      }
    }
  }

  const isScanned = (totalTextChars / pagesToCheck) < 30 && pdf.numPages > 0;
  
  return { count: totalImages, isScanned };
}

/**
 * 將 PDF 內嵌圖片數據轉換為 Blob (PNG)
 */
async function imageToBlob(img: any): Promise<Blob | null> {
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  if (!ctx) return null;

  try {
    // 處理 ImageBitmap (現代瀏覽器中 PDF.js 偏好的格式)
    if (img.bitmap && img.bitmap instanceof ImageBitmap) {
      canvas.width = img.bitmap.width;
      canvas.height = img.bitmap.height;
      ctx.drawImage(img.bitmap, 0, 0);
    } 
    // 處理原始像素數據
    else if (img.data) {
      canvas.width = img.width;
      canvas.height = img.height;
      const imageData = ctx.createImageData(img.width, img.height);
      
      // 如果是 RGB 數據 (長度為 w * h * 3)，需要填充 Alpha 通道
      if (img.data.length === img.width * img.height * 3) {
        for (let i = 0, j = 0; i < img.data.length; i += 3, j += 4) {
          imageData.data[j] = img.data[i];
          imageData.data[j+1] = img.data[i+1];
          imageData.data[j+2] = img.data[i+2];
          imageData.data[j+3] = 255;
        }
      } else {
        // 假設已經是 RGBA
        imageData.data.set(img.data);
      }
      ctx.putImageData(imageData, 0, 0);
    } else {
      return null;
    }

    const blob = await new Promise<Blob | null>(resolve => canvas.toBlob(resolve, 'image/png'));
    
    // 記憶體優化：清空 canvas 尺寸有助於某些瀏覽器釋放後端資源
    canvas.width = 0;
    canvas.height = 0;
    
    return blob;
  } catch (e) {
    console.error('Image to Blob conversion failed:', e);
    return null;
  }
}
