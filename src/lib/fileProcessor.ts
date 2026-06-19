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

// Helper to sanitize filename
function sanitizeFilename(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, '_');
}

// Helper to check if a file path has an image extension
function isImageFile(path: string): boolean {
  const ext = path.split('.').pop()?.toLowerCase();
  return !!(ext && ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'emf', 'wmf', 'svg', 'tiff', 'tif', 'webp'].includes(ext));
}

// Helper to validate if a blob is a valid displayable image in browser
async function isValidBrowserImage(blob: Blob): Promise<boolean> {
  // If the blob size is extremely tiny (e.g., less than 100 bytes), it is likely a bad or empty placeholder
  if (blob.size < 100) return false;
  
  return new Promise((resolve) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    
    let timer = setTimeout(() => {
      img.onload = null;
      img.onerror = null;
      URL.revokeObjectURL(url);
      resolve(false); // timeout, invalid/unsupported image
    }, 2000); // 2 seconds timeout for safety
    
    img.onload = () => {
      clearTimeout(timer);
      URL.revokeObjectURL(url);
      resolve(img.width > 0 && img.height > 0);
    };
    img.onerror = () => {
      clearTimeout(timer);
      URL.revokeObjectURL(url);
      resolve(false);
    };
    img.src = url;
  });
}

// Helper to normalized OOXML relationship paths
function normalizeExcelPath(baseDir: string, target: string): string {
  if (target.startsWith('/')) {
    return target.substring(1);
  }
  const parts = baseDir.split('/').filter(Boolean);
  const targetParts = target.split('/');
  for (const p of targetParts) {
    if (p === '.') {
      continue;
    } else if (p === '..') {
      parts.pop();
    } else {
      parts.push(p);
    }
  }
  return parts.join('/');
}

// Helper to look up element attributes in a namespace insensitive way
const getAttrVal = (el: Element, attrName: string): string | null => {
  const exact = el.getAttribute(attrName);
  if (exact !== null) return exact;
  const lowercase = el.getAttribute(attrName.toLowerCase());
  if (lowercase !== null) return lowercase;
  
  for (let i = 0; i < el.attributes.length; i++) {
    const attr = el.attributes[i];
    if (attr.localName === attrName || attr.localName?.toLowerCase() === attrName.toLowerCase()) {
      return attr.value;
    }
  }
  return null;
};

// Helper lookup for any elements matching the localName across DOM
const getElementsByLocalName = (parent: Document | Element, localName: string): Element[] => {
  const all = parent.getElementsByTagName('*');
  const result: Element[] = [];
  for (let i = 0; i < all.length; i++) {
    if (all[i].localName === localName) {
      result.push(all[i]);
    }
  }
  return result;
};

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
  
  interface ExtractionItem {
    mediaPath: string;
    sheetName?: string;
    pageNumber?: number;
    row?: number;
    col?: number;
  }
  
  const finalExtractions: ExtractionItem[] = [];
  
  // 計算本 Office 的媒體資料夾基礎路徑
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
  
  // 彙整 ZIP 中所有媒體檔案，確保 100% 都不缺失
  const allMediaFilesInZip = Object.keys(zip.files).filter(path => path.startsWith(mediaPathPrefix) && !zip.files[path].dir && isImageFile(path));
  
  if (isXlsx) {
    try {
      // 1. 解析 xl/workbook.xml 及其 rels，建立工作表檔案路徑與真實工作表名稱的映射
      const sheetPathToName = new Map<string, string>();
      try {
        const workbookXmlText = await zip.file('xl/workbook.xml')?.async('string');
        const workbookRelsText = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
        
        if (workbookXmlText && workbookRelsText) {
          const workbookDoc = parser.parseFromString(workbookXmlText, 'application/xml');
          const relsDoc = parser.parseFromString(workbookRelsText, 'application/xml');
          
          const rIdToTarget = new Map<string, string>();
          const relationships = getElementsByLocalName(relsDoc, 'Relationship');
          for (const rel of relationships) {
            const id = getAttrVal(rel, 'Id');
            const target = getAttrVal(rel, 'Target');
            if (id && target) {
              const fullTarget = normalizeExcelPath('xl', target);
              rIdToTarget.set(id, fullTarget);
            }
          }
          
          const sheets = getElementsByLocalName(workbookDoc, 'sheet');
          for (const sheet of sheets) {
            const name = getAttrVal(sheet, 'name') || '';
            const rId = getAttrVal(sheet, 'id') || getAttrVal(sheet, 'r:id');
            if (name && rId) {
              const targetPath = rIdToTarget.get(rId);
              if (targetPath) {
                sheetPathToName.set(targetPath, name);
              }
            }
          }
        }
      } catch (e) {
        console.warn('解析 Excel 工作表名稱失敗:', e);
      }
      
      // 2. 依照工作表順序，查找關聯的 Drawing 檔案 (sheet rels)
      const drawingToSheetName = new Map<string, string>();
      const drawingOrder: string[] = [];
      
      const sheetFiles = Object.keys(zip.files).filter(path => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path));
      const orderedSheets = Array.from(sheetPathToName.keys()).filter(path => zip.file(path));
      for (const sf of sheetFiles) {
        if (!orderedSheets.includes(sf)) {
          orderedSheets.push(sf);
        }
      }
      
      for (const sheetPath of orderedSheets) {
        const sheetXmlName = sheetPath.split('/').pop()!;
        const relsPath = `xl/worksheets/_rels/${sheetXmlName}.rels`;
        const resolvedSheetName = sheetPathToName.get(sheetPath) || sheetXmlName.replace('.xml', '');
        
        if (zip.file(relsPath)) {
          const relsText = await zip.file(relsPath)!.async('string');
          const relsDoc = parser.parseFromString(relsText, 'application/xml');
          const relationships = getElementsByLocalName(relsDoc, 'Relationship');
          for (const rel of relationships) {
            const type = getAttrVal(rel, 'Type');
            const target = getAttrVal(rel, 'Target');
            if (type?.endsWith('/drawing') && target) {
              const resolvedTarget = normalizeExcelPath('xl/worksheets', target);
              if (!drawingOrder.includes(resolvedTarget)) {
                drawingOrder.push(resolvedTarget);
              }
              drawingToSheetName.set(resolvedTarget, resolvedSheetName);
            }
          }
        }
      }
      
      // 3. 備用方案：追加其餘 Drawings 項目
      const allDrawings = Object.keys(zip.files).filter(path => /^xl\/drawings\/drawing\d+\.xml$/i.test(path));
      allDrawings.sort((a, b) => {
        const numA = parseInt(a.replace(/\D/g, '') || '0', 10);
        const numB = parseInt(b.replace(/\D/g, '') || '0', 10);
        return numA - numB;
      });
      for (const dr of allDrawings) {
        if (!drawingOrder.includes(dr)) {
          drawingOrder.push(dr);
          if (!drawingToSheetName.has(dr)) {
            drawingToSheetName.set(dr, '其他');
          }
        }
      }
      
      const collectedImages: {
        drawingIndex: number;
        sheetName: string;
        row: number;
        col: number;
        mediaPath: string;
      }[] = [];
      
      const findInSubtree = (parent: Element, name: string): Element | null => {
        const elements = parent.getElementsByTagName('*');
        for (let i = 0; i < elements.length; i++) {
          if (elements[i].localName === name) {
            return elements[i];
          }
        }
        return null;
      };
      
      // 4. 解析各繪圖 XML 中形狀/圖片的儲存格坐標 (Row & Col)
      for (let d = 0; d < drawingOrder.length; d++) {
        const drawingPath = drawingOrder[d];
        const drawingXmlFile = zip.file(drawingPath);
        if (!drawingXmlFile) continue;
        
        const drawingText = await drawingXmlFile.async('string');
        const lastSlash = drawingPath.lastIndexOf('/');
        const dir = drawingPath.substring(0, lastSlash);
        const fileXmlName = drawingPath.substring(lastSlash + 1);
        const relsPath = `${dir}/_rels/${fileXmlName}.rels`;
        const currentSheetName = drawingToSheetName.get(drawingPath) || '其他';
        
        const rIdToMedia = new Map<string, string>();
        if (zip.file(relsPath)) {
          const relsText = await zip.file(relsPath)!.async('string');
          const relsDoc = parser.parseFromString(relsText, 'application/xml');
          const relationships = getElementsByLocalName(relsDoc, 'Relationship');
          for (const rel of relationships) {
            const id = getAttrVal(rel, 'Id');
            const target = getAttrVal(rel, 'Target');
            if (id && target) {
              const fullTarget = normalizeExcelPath(dir, target);
              if (isImageFile(fullTarget)) {
                rIdToMedia.set(id, fullTarget);
              }
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
            rId = getAttrVal(blipNode, 'embed');
          }
          
          if (rId) {
            const mediaPath = rIdToMedia.get(rId);
            if (mediaPath && zip.file(mediaPath)) {
              collectedImages.push({
                drawingIndex: d,
                sheetName: currentSheetName,
                row,
                col,
                mediaPath
              });
            }
          }
        }
      }
      
      // 5. 排序：工作表(繪圖索引)優先 -> 橫列(Row)優先 -> 直欄(Col)
      collectedImages.sort((a, b) => {
        if (a.drawingIndex !== b.drawingIndex) {
          return a.drawingIndex - b.drawingIndex;
        }
        if (a.row !== b.row) {
          return a.row - b.row;
        }
        return a.col - b.col;
      });
      
      for (const img of collectedImages) {
        finalExtractions.push({
          mediaPath: img.mediaPath,
          sheetName: img.sheetName,
          row: img.row,
          col: img.col
        });
      }
      
      // 6. 補上其他未在 Drawing 被參照的媒體檔案 (如頁首、浮水印等)
      const extractedPaths = new Set(collectedImages.map(img => img.mediaPath));
      for (const path of allMediaFilesInZip) {
        if (!extractedPaths.has(path)) {
          finalExtractions.push({
            mediaPath: path,
            sheetName: '其他'
          });
        }
      }
    } catch (e) {
      console.warn('XLSX 工作表與儲存格排序解析失敗，退回傳統模式', e);
      for (const path of allMediaFilesInZip) {
        finalExtractions.push({ mediaPath: path, sheetName: '工作表' });
      }
    }
  } else if (isDocx) {
    try {
      // 1. 載入 Word 關聯對應
      const rIdToMedia = new Map<string, string>();
      const relsPath = 'word/_rels/document.xml.rels';
      if (zip.file(relsPath)) {
        const relsText = await zip.file(relsPath)!.async('string');
        const relsDoc = parser.parseFromString(relsText, 'application/xml');
        const relationships = getElementsByLocalName(relsDoc, 'Relationship');
        for (const rel of relationships) {
          const id = getAttrVal(rel, 'Id');
          const target = getAttrVal(rel, 'Target');
          if (id && target) {
            let fullTarget = target;
            if (!target.startsWith('word/')) {
              fullTarget = 'word/' + target;
            }
            if (isImageFile(fullTarget)) {
              rIdToMedia.set(id, fullTarget);
            }
          }
        }
      }
      
      // 2. 遞迴走訪主文檔，估計分頁與保留物理流順序
      const docPath = 'word/document.xml';
      if (zip.file(docPath)) {
        const docText = await zip.file(docPath)!.async('string');
        const doc = parser.parseFromString(docText, 'application/xml');
        
        let currentPage = 1;
        
        const walkNode = (node: Node) => {
          if (node.nodeType === Node.ELEMENT_NODE) {
            const el = node as Element;
            const localName = el.localName;
            
            if (localName === 'lastRenderedPageBreak') {
              currentPage++;
            } else if (localName === 'br' && getAttrVal(el, 'type') === 'page') {
              currentPage++;
            }
            
            let rId = getAttrVal(el, 'embed');
            if (!rId) {
              rId = getAttrVal(el, 'id');
            }
            
            if (rId && rIdToMedia.has(rId)) {
              const mediaPath = rIdToMedia.get(rId)!;
              if (zip.file(mediaPath)) {
                finalExtractions.push({
                  mediaPath,
                  pageNumber: currentPage
                });
              }
            }
          }
          
          let child = node.firstChild;
          while (child) {
            walkNode(child);
            child = child.nextSibling;
          }
        };
        
        walkNode(doc);
      }
      
      // 3. 補上其他未在主流程中抓到的圖片 (背景、頁碼、備份)
      const extractedPaths = new Set(finalExtractions.map(ext => ext.mediaPath));
      for (const path of allMediaFilesInZip) {
        if (!extractedPaths.has(path)) {
          finalExtractions.push({
            mediaPath: path,
            pageNumber: 1
          });
        }
      }
    } catch (e) {
      console.warn('DOCX 走訪解析失敗，退回傳統分頁模式', e);
      for (const path of allMediaFilesInZip) {
        finalExtractions.push({ mediaPath: path, pageNumber: 1 });
      }
    }
  } else {
    // 其他 Office 文件格式直接順序解析
    for (const path of allMediaFilesInZip) {
      finalExtractions.push({ mediaPath: path });
    }
  }
  
  // 建立各 Sheet / 各 Page 子獨立增序列
  const excelSheetCounters = new Map<string, number>();
  const docxPageCounters = new Map<number, number>();
  
  let count = 0;
  for (let i = 0; i < finalExtractions.length; i++) {
    const item = finalExtractions[i];
    try {
      const zipFile = zip.files[item.mediaPath];
      if (!zipFile || zipFile.dir) continue;
      
      const blob = await zipFile.async('blob');
      
      // Filter out invalid, zero-size, or non-displayable formats like EMF/WMF which browser cannot render
      const isValid = await isValidBrowserImage(blob);
      if (!isValid) {
        continue;
      }
      
      let fileName = '';
      const originalName = item.mediaPath.split('/').pop() || `image_${count + 1}`;
      const ext = originalName.split('.').pop()?.toLowerCase() || 'png';
      
      if (isXlsx) {
        // Excel 命名格式：[工作表名稱]_image[3位補零工作表記數].[副檔名]
        const sheetName = item.sheetName || '工作表';
        const nextIdx = (excelSheetCounters.get(sheetName) || 0) + 1;
        excelSheetCounters.set(sheetName, nextIdx);
        
        const prefix = String(nextIdx).padStart(3, '0');
        const sanitizedSheet = sanitizeFilename(sheetName);
        fileName = `${sanitizedSheet}_image${prefix}.${ext}`;
      } else if (isDocx) {
        // Word 命名格式：page[頁碼]_image[3位補零頁計數].[副檔名]
        const pageNum = item.pageNumber || 1;
        const nextIdx = (docxPageCounters.get(pageNum) || 0) + 1;
        docxPageCounters.set(pageNum, nextIdx);
        
        const prefix = String(nextIdx).padStart(3, '0');
        fileName = `page${pageNum}_image${prefix}.${ext}`;
      } else {
        // 其他文件
        const prefix = String(count + 1).padStart(3, '0');
        fileName = `image${prefix}.${ext}`;
      }
      
      const fileHandle = await targetDirHandle.getFileHandle(fileName, { create: true });
      const writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();
      count++;
    } catch (e) {
      console.warn(`跳過特定圖片擷取，路徑: ${item.mediaPath}`, e);
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
  
  // 計算 1/3 的總頁數作為檢測斷點 (最少要檢測 1 頁)
  const checkPointPage = Math.max(1, Math.ceil(pdf.numPages / 3));
  let isScanned = false;
  const pdfPageCounters = new Map<number, number>();
  
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    
    // 執行檔案三分之一內容，收集文字資訊
    if (i <= checkPointPage) {
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
              const nextIdx = (pdfPageCounters.get(i) || 0) + 1;
              pdfPageCounters.set(i, nextIdx);
              const prefix = String(nextIdx).padStart(3, '0');
              const fileName = `page${i}_image${prefix}.png`;
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

    // 達到 1/3 頁數時進行判定：若這 1/3 內容沒偵測到文字，則判定為掃描檔並中斷執行
    if (i === checkPointPage) {
      if (totalTextChars < 30 && pdf.numPages > 0) {
        isScanned = true;
        console.log(`PDF 偵測結果: 在前三分之一內容 (${checkPointPage}/${pdf.numPages} 頁) 中僅偵測到 ${totalTextChars} 個字元，判定為掃描型 PDF，中止擷取。`);
        break; // 中止後續頁面處理
      } else {
        console.log(`PDF 偵測結果: 在前三分之一內容 (${checkPointPage}/${pdf.numPages} 頁) 中偵測到 ${totalTextChars} 個字元，保留文字型 PDF，持續執行。`);
      }
    }
  }
  
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
