import JSZip from 'jszip';
import * as pdfjsLib from 'pdfjs-dist';
import * as CFB from 'cfb';

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

// Helper to get file base name without extension, and sanitized
function getFileBaseName(filename: string): string {
  const lastDot = filename.lastIndexOf('.');
  const base = lastDot === -1 ? filename : filename.substring(0, lastDot);
  return sanitizeFilename(base);
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

/**
 * 安全地將 Blob 寫入指定目錄下的檔案，內建強大的重試機制以避開 Chromium File System Access API
 * 的 "An operation that depends on state cached in an interface object..." 快取/鎖定衝突錯誤。
 */
async function safeWriteFile(
  dirHandle: FileSystemDirectoryHandle,
  fileName: string,
  blob: Blob
): Promise<boolean> {
  const maxAttempts = 5;
  let attempts = 0;
  
  while (attempts < maxAttempts) {
    let fileHandle: FileSystemFileHandle | null = null;
    let writable: FileSystemWritableFileStream | null = null;
    
    try {
      fileHandle = await dirHandle.getFileHandle(fileName, { create: true });
      writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();
      return true; // 成功寫入
    } catch (e: any) {
      attempts++;
      console.warn(
        `寫入檔案 ${fileName} 失敗 (嘗試 ${attempts}/${maxAttempts}):`,
        e?.message || e
      );
      
      // 確保在失敗時嘗試關閉 writable 以免洩漏資源
      if (writable) {
        try {
          await writable.close();
        } catch (_) {}
      }
      
      if (attempts >= maxAttempts) {
        throw e; // 如果重試都失敗，則拋出原本的錯誤
      }
      
      // 使用遞增延遲 (50ms, 100ms, 200ms, 400ms...) 讓作業系統與瀏覽器線程有時間釋放檔案與更新快取
      const delay = 50 * Math.pow(2, attempts - 1);
      await new Promise(resolve => setTimeout(resolve, delay));
    }
  }
  return false;
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

// Helper to fetch files from JSZip in a case-insensitive manner
function zipGetFile(zip: JSZip, path: string): JSZip.JSZipObject | null {
  const lowerPath = path.toLowerCase();
  const foundKey = Object.keys(zip.files).find(k => k.toLowerCase() === lowerPath);
  return foundKey ? zip.file(foundKey) : null;
}

// Counts references of media files across all OOXML .rels files to handle duplicates accurately in fallback mode
async function countMediaReferences(zip: JSZip, allMedia: string[]): Promise<Map<string, number>> {
  const refCounts = new Map<string, number>();
  for (const path of allMedia) {
    refCounts.set(path, 0); // Initialize reference counts
  }
  
  const parser = new DOMParser();
  const relsFiles = Object.keys(zip.files).filter(p => p.toLowerCase().endsWith('.rels'));
  
  for (const relsPath of relsFiles) {
    try {
      const relsObj = zip.files[relsPath];
      if (!relsObj || relsObj.dir) continue;
      
      const text = await relsObj.async('string');
      const doc = parser.parseFromString(text, 'application/xml');
      const relationships = getElementsByLocalName(doc, 'Relationship');
      
      const lastSlash = relsPath.lastIndexOf('/');
      let dir = lastSlash === -1 ? '' : relsPath.substring(0, lastSlash);
      if (dir.toLowerCase().endsWith('/_rels')) {
        dir = dir.substring(0, dir.length - 6);
      } else if (dir.toLowerCase() === '_rels') {
        dir = '';
      }
      
      for (const rel of relationships) {
        const target = getAttrVal(rel, 'Target');
        if (target) {
          const fullTarget = normalizeExcelPath(dir, target);
          const foundKey = allMedia.find(m => m.toLowerCase() === fullTarget.toLowerCase());
          if (foundKey) {
            refCounts.set(foundKey, (refCounts.get(foundKey) || 0) + 1);
          }
        }
      }
    } catch (e) {
      console.warn(`解析關聯檔案 ${relsPath} 失敗:`, e);
    }
  }
  return refCounts;
}

// Parses worksheet sheet names from legacy XLS BOUNDSHEET records (0x0085)
export function parseXlsSheetNames(workbookBytes: Uint8Array): string[] {
  const sheetNames: string[] = [];
  let offset = 0;
  const len = workbookBytes.length;
  
  while (offset < len - 4) {
    const type = workbookBytes[offset] | (workbookBytes[offset + 1] << 8);
    const recordLength = workbookBytes[offset + 2] | (workbookBytes[offset + 3] << 8);
    
    if (type === 0x0085) { // BOUNDSHEET record
      try {
        const dataStart = offset + 4;
        const nameLen = workbookBytes[dataStart + 6];
        const isUnicode = workbookBytes[dataStart + 7]; // 0 = compressed ASCII, 1 = uncompressed UTF-16
        
        let name = '';
        if (isUnicode === 1) {
          for (let k = 0; k < nameLen; k++) {
            const charCode = workbookBytes[dataStart + 8 + k * 2] | (workbookBytes[dataStart + 9 + k * 2] << 8);
            name += String.fromCharCode(charCode);
          }
        } else {
          for (let k = 0; k < nameLen; k++) {
            name += String.fromCharCode(workbookBytes[dataStart + 8 + k]);
          }
        }
        if (name) {
          sheetNames.push(name);
        }
      } catch (e) {
        console.warn('解析 XLS BOUNDSHEET 發生錯誤:', e);
      }
    }
    
    offset += 4 + recordLength;
  }
  return sheetNames;
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
    const firstMediaFolder = Object.keys(zip.files).find(p => p.toLowerCase().includes('/media/'));
    if (firstMediaFolder) {
      const idx = firstMediaFolder.toLowerCase().indexOf('/media/');
      mediaPathPrefix = firstMediaFolder.substring(0, idx + 7);
    }
  }
  
  // 彙整 ZIP 中所有媒體檔案，確保 100% 都不缺失
  const allMediaFilesInZip = Object.keys(zip.files).filter(path => path.toLowerCase().startsWith(mediaPathPrefix.toLowerCase()) && !zip.files[path].dir && isImageFile(path));
  
  // 異步解析關聯統計，做為後續重複提取機制的兜底
  const mediaRefCounts = await countMediaReferences(zip, allMediaFilesInZip);

  if (isXlsx) {
    try {
      // 1. 解析 xl/workbook.xml 及其 rels，建立工作表檔案路徑與真實工作表名稱的映射
      const sheetPathToName = new Map<string, string>();
      try {
        const workbookXmlFile = zipGetFile(zip, 'xl/workbook.xml');
        const workbookRelsFile = zipGetFile(zip, 'xl/_rels/workbook.xml.rels');
        
        if (workbookXmlFile && workbookRelsFile) {
          const workbookXmlText = await workbookXmlFile.async('string');
          const workbookRelsText = await workbookRelsFile.async('string');
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
      const orderedSheets = Array.from(sheetPathToName.keys()).filter(path => zipGetFile(zip, path));
      for (const sf of sheetFiles) {
        if (!orderedSheets.includes(sf)) {
          orderedSheets.push(sf);
        }
      }
      
      for (const sheetPath of orderedSheets) {
        try {
          const sheetXmlName = sheetPath.split('/').pop()!;
          const relsPath = `xl/worksheets/_rels/${sheetXmlName}.rels`;
          const resolvedSheetName = sheetPathToName.get(sheetPath) || sheetXmlName.replace('.xml', '');
          
          const relsFile = zipGetFile(zip, relsPath);
          if (relsFile) {
            const relsText = await relsFile.async('string');
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
        } catch (e) {
          console.warn(`解析工作表關聯檔案 ${sheetPath} 失敗:`, e);
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
        try {
          const drawingXmlFile = zipGetFile(zip, drawingPath);
          if (!drawingXmlFile) continue;
          
          const drawingText = await drawingXmlFile.async('string');
          const lastSlash = drawingPath.lastIndexOf('/');
          const dir = drawingPath.substring(0, lastSlash);
          const fileXmlName = drawingPath.substring(lastSlash + 1);
          const relsPath = `${dir}/_rels/${fileXmlName}.rels`;
          const currentSheetName = drawingToSheetName.get(drawingPath) || '其他';
          
          const rIdToMedia = new Map<string, string>();
          const relsFile = zipGetFile(zip, relsPath);
          if (relsFile) {
            const relsText = await relsFile.async('string');
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
              rId = getAttrVal(blipNode, 'embed') || getAttrVal(blipNode, 'link') || getAttrVal(blipNode, 'id');
            }
            
            if (rId) {
              const mediaPath = rIdToMedia.get(rId);
              if (mediaPath && zipGetFile(zip, mediaPath)) {
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
        } catch (e) {
          console.warn(`解析繪圖檔案 ${drawingPath} 失敗:`, e);
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
          // 在兜底時，考慮此圖片在 .rels 中實際被引用的總次數
          const timesReferenced = Math.max(1, mediaRefCounts.get(path) || 0);
          for (let r = 0; r < timesReferenced; r++) {
            finalExtractions.push({
              mediaPath: path,
              sheetName: '其他'
            });
          }
        }
      }
    } catch (e) {
      console.warn('XLSX 工作表與儲存格排序解析失敗，退回傳統模式', e);
      for (const path of allMediaFilesInZip) {
        const timesReferenced = Math.max(1, mediaRefCounts.get(path) || 0);
        for (let r = 0; r < timesReferenced; r++) {
          finalExtractions.push({ mediaPath: path, sheetName: '工作表' });
        }
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
      const baseName = getFileBaseName(file.name);
      
      if (isXlsx) {
        // Excel 命名格式：[工作表名稱]_image[3位補零工作表記數].[副檔名]
        const sheetName = item.sheetName || '工作表';
        const nextIdx = (excelSheetCounters.get(sheetName) || 0) + 1;
        excelSheetCounters.set(sheetName, nextIdx);
        
        const prefix = String(nextIdx).padStart(3, '0');
        const sanitizedSheet = sanitizeFilename(sheetName);
        fileName = `${baseName}_${sanitizedSheet}_image${prefix}.${ext}`;
      } else if (isDocx) {
        // Word 命名格式：page[頁碼]_image[3位補零頁計數].[副檔名]
        const pageNum = item.pageNumber || 1;
        const nextIdx = (docxPageCounters.get(pageNum) || 0) + 1;
        docxPageCounters.set(pageNum, nextIdx);
        
        const prefix = String(nextIdx).padStart(3, '0');
        fileName = `${baseName}_page${pageNum}_image${prefix}.${ext}`;
      } else {
        // 其他文件
        const prefix = String(count + 1).padStart(3, '0');
        fileName = `${baseName}_image${prefix}.${ext}`;
      }
      
      await safeWriteFile(targetDirHandle, fileName, blob);
      count++;
    } catch (e) {
      console.warn(`跳過特定圖片擷取，路徑: ${item.mediaPath}`, e);
    }
  }
  
  return count;
}

export async function extractImagesFromPdf(file: File, targetDirHandle: FileSystemDirectoryHandle): Promise<{ count: number; isScanned: boolean }> {
  const arrayBuffer = await file.arrayBuffer();
  const baseName = getFileBaseName(file.name);
  
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
    
    // 收集每一頁的文字資訊
    const textContent = await page.getTextContent();
    const text = textContent.items.map((item: any) => item.str).join('');
    const pageTextLength = text.trim().length;
    
    if (i <= checkPointPage) {
      totalTextChars += pageTextLength;
    }
    
    // 達到 1/3 頁數時進行判定：若這 1/3 內容沒偵測到任何文字 (totalTextChars === 0)，則判定為掃描檔並中斷執行
    if (i === checkPointPage) {
      if (totalTextChars === 0 && pdf.numPages > 0) {
        isScanned = true;
        console.log(`PDF 偵測結果: 在前三分之一內容 (${checkPointPage}/${pdf.numPages} 頁) 中未偵測到任何文字，判定為掃描型 PDF，中止擷取。`);
        break; // 中止後續頁面處理
      } else {
        console.log(`PDF 偵測結果: 在前三分之一內容 (${checkPointPage}/${pdf.numPages} 頁) 中偵測到 ${totalTextChars} 個字元，保留文字型 PDF，持續執行。`);
      }
    }
    
    // 如果單頁沒有文字 (也就是整頁掃描圖或無文字頁)，則跳過該頁的所有圖片擷取，避免擷取到整頁掃描
    if (pageTextLength === 0) {
      console.log(`第 ${i} 頁無偵測到文字，跳過該頁圖片擷取 (避免擷取整頁掃描圖)`);
      continue;
    }
    
    const operatorList = await page.getOperatorList();
    const validImages = [
      pdfjsLib.OPS.paintImageXObject,
      pdfjsLib.OPS.paintInlineImageXObject,
      pdfjsLib.OPS.paintImageXObjectRepeat,
      pdfjsLib.OPS.paintImageMaskXObject
    ].filter(Boolean);

    for (let j = 0; j < operatorList.fnArray.length; j++) {
      if (validImages.includes(operatorList.fnArray[j])) {
        const arg = operatorList.argsArray[j][0];
        
        try {
          let img: any = null;
          if (operatorList.fnArray[j] === pdfjsLib.OPS.paintInlineImageXObject) {
            // 如果是內嵌圖片 (Inline Image)，直接將對象做為 img (它已經是圖片對象本身)
            img = arg;
          } else {
            // 傳統 XObject，其第一引數為物件 ID 字串
            const objId = arg;
            
            // 使用 Promise 配合 PDFObjects 內部回呼機制，確保非同步載入的圖片在 objs/commonObjs 完全 ready 後再讀取
            img = await new Promise<any>((resolve) => {
              let resolved = false;
              
              // 優先由 page.objs 取得
              const res = page.objs.get(objId, (data: any) => {
                if (!resolved) {
                  resolved = true;
                  resolve(data);
                }
              });
              if (res) {
                resolved = true;
                resolve(res);
                return;
              }
              
              // 其次嘗試 page.commonObjs
              const commonRes = page.commonObjs.get(objId, (data: any) => {
                if (!resolved) {
                  resolved = true;
                  resolve(data);
                }
              });
              if (commonRes) {
                resolved = true;
                resolve(commonRes);
                return;
              }
              
              // 設定 2 秒逾時兜底，避免圖片毀損或遺失導致執行程序掛起
              setTimeout(() => {
                if (!resolved) {
                  resolved = true;
                  resolve(null);
                }
              }, 2000);
            });
          }

          if (img && (img.data || img.bitmap)) {
            const blob = await imageToBlob(img);
            if (blob) {
              const nextIdx = (pdfPageCounters.get(i) || 0) + 1;
              pdfPageCounters.set(i, nextIdx);
              const prefix = String(nextIdx).padStart(3, '0');
              const fileName = `${baseName}_page${i}_image${prefix}.png`;
              await safeWriteFile(targetDirHandle, fileName, blob);
              totalImages++;
            }
          }
        } catch (e) {
          console.warn(`第 ${i} 頁圖片擷取失敗:`, e);
        }
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
      
      const pixelCount = img.width * img.height;
      if (img.data.length === pixelCount) {
        // Greyscale (1 channel)
        for (let i = 0, j = 0; i < img.data.length && j < imageData.data.length; i++, j += 4) {
          const val = img.data[i];
          imageData.data[j] = val;
          imageData.data[j+1] = val;
          imageData.data[j+2] = val;
          imageData.data[j+3] = 255;
        }
      } else if (img.data.length === pixelCount * 2) {
        // Greyscale + Alpha (2 channels)
        for (let i = 0, j = 0; i < img.data.length && j < imageData.data.length; i += 2, j += 4) {
          const val = img.data[i];
          imageData.data[j] = val;
          imageData.data[j+1] = val;
          imageData.data[j+2] = val;
          imageData.data[j+3] = img.data[i+1];
        }
      } else if (img.data.length === pixelCount * 3) {
        // RGB (3 channels)
        for (let i = 0, j = 0; i < img.data.length && j < imageData.data.length; i += 3, j += 4) {
          imageData.data[j] = img.data[i];
          imageData.data[j+1] = img.data[i+1];
          imageData.data[j+2] = img.data[i+2];
          imageData.data[j+3] = 255;
        }
      } else {
        // Safe fallback copy
        const copyLen = Math.min(img.data.length, imageData.data.length);
        for (let i = 0; i < copyLen; i++) {
          imageData.data[i] = img.data[i];
        }
        // Fill remaining alphas with 255 if needed
        if (img.data.length < imageData.data.length) {
          for (let i = 3; i < imageData.data.length; i += 4) {
            if (imageData.data[i] === 0 && (i < img.data.length ? img.data[i] === undefined : true)) {
              imageData.data[i] = 255;
            }
          }
        }
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

/**
 * 從 XLS 檔案中提取 Workbook (或 Book) 二進位資料流
 */
function getXlsWorkbookStream(arrayBuffer: ArrayBuffer): Uint8Array | null {
  try {
    const cfbObj = (CFB as any).default || CFB;
    const container = cfbObj.read(new Uint8Array(arrayBuffer), { type: 'array' });
    const fileIndex = container.FileIndex;
    const entry = fileIndex.find((f: any) => {
      const n = f.name || '';
      return n === 'Workbook' || n === 'Book' || n === 'WORKBOOK' || n === 'BOOK';
    });
    if (entry && entry.content) {
      return new Uint8Array(entry.content);
    }
  } catch (e) {
    console.error('解析 XLS (OLE CFB 容器) 失敗:', e);
  }
  return null;
}

/**
 * 重組 XLS 中「所有」被 CONTINUE 紀錄中斷的資料流，並拼裝成一個完全去除了 BIFF8 紀錄與分段包頭的純淨 Payload 串流。
 * 這能確保無論圖片位於哪個 Sheet (分頁) 的哪個位置，都能被無縫還原，並完全排除 0x003C (CONTINUE) 紀錄頭造成的解碼花屏。
 */
function reassembleXlsDrawingStream(workbookBytes: Uint8Array): Uint8Array {
  const records: { type: number; data: Uint8Array }[] = [];
  
  let offset = 0;
  const len = workbookBytes.length;
  
  while (offset < len - 4) {
    const type = workbookBytes[offset] | (workbookBytes[offset + 1] << 8);
    const recordLength = workbookBytes[offset + 2] | (workbookBytes[offset + 3] << 8);
    
    const dataStart = offset + 4;
    const dataEnd = Math.min(len, dataStart + recordLength);
    const data = workbookBytes.subarray(dataStart, dataEnd);
    
    if (type === 0x003C) { // CONTINUE 紀錄
      // 如果前一個紀錄存在，就把這個 CONTINUE 的數據拼接到前一個紀錄中
      if (records.length > 0) {
        const prev = records[records.length - 1];
        const combined = new Uint8Array(prev.data.length + data.length);
        combined.set(prev.data);
        combined.set(data, prev.data.length);
        prev.data = combined;
      } else {
        records.push({ type, data });
      }
    } else {
      records.push({ type, data });
    }
    
    offset = dataEnd;
  }
  
  // 組合所有無縫還原後的紀錄 Payload 為單一連續的 Uint8Array 供雕刻使用
  let totalLen = 0;
  for (const r of records) {
    totalLen += r.data.length;
  }
  
  const combined = new Uint8Array(totalLen);
  let writeOffset = 0;
  for (const r of records) {
    combined.set(r.data, writeOffset);
    writeOffset += r.data.length;
  }
  
  return combined;
}

/**
 * 依據 JPEG 區段結構解析 true EOI 結尾，完全跳過內嵌的 EXIF 縮圖與 false positive 的 FF D9
 */
function findJpegEnd(bytes: Uint8Array, startOffset: number): number {
  let offset = startOffset;
  const len = bytes.length;
  
  if (bytes[offset] !== 0xFF || bytes[offset + 1] !== 0xD8) {
    return -1;
  }
  
  offset += 2;
  
  while (offset < len - 4) {
    if (bytes[offset] !== 0xFF) {
      // 搜尋下一個 0xFF
      const nextFF = bytes.indexOf(0xFF, offset);
      if (nextFF === -1 || nextFF > len - 4) {
        return -1;
      }
      offset = nextFF;
    }
    
    const marker = bytes[offset + 1];
    
    if (marker === 0xFF) {
      offset++;
      continue;
    }
    
    if (marker === 0xD9) {
      return offset + 2;
    }
    
    if (marker >= 0xD0 && marker <= 0xD7) {
      offset += 2;
      continue;
    }
    
    if (marker === 0xDA) {
      if (offset + 3 >= len) return -1;
      // SOS (Start of Scan) 進入壓縮編碼數據區，此時必須掃描直到真正的 FF D9
      const sosHeaderLen = (bytes[offset + 2] << 8) | bytes[offset + 3];
      offset += 2 + sosHeaderLen;
      
      while (offset < len - 1) {
        const nextFF = bytes.indexOf(0xFF, offset);
        if (nextFF === -1 || nextFF >= len - 1) {
          return -1;
        }
        
        const nextMarker = bytes[nextFF + 1];
        if (nextMarker === 0x00 || (nextMarker >= 0xD0 && nextMarker <= 0xD7)) {
          // 跳過填充 of 0x00 或 Restart 標記
          offset = nextFF + 2;
        } else if (nextMarker === 0xD9) {
          return nextFF + 2;
        } else {
          offset = nextFF;
          break;
        }
      }
      continue;
    }
    
    if (offset + 3 >= len) return -1;
    const markerLen = (bytes[offset + 2] << 8) | bytes[offset + 3];
    if (markerLen < 2) {
      offset += 2;
    } else {
      offset += 2 + markerLen;
    }
  }
  
  return -1;
}

/**
 * 依據 PNG 區段結構解析 IEND 結尾，確保 100% 精準取得圖片二進位大小，完全避免 false positive IEND 的匹配與截斷。
 */
function findPngEnd(bytes: Uint8Array, startOffset: number): number {
  let offset = startOffset;
  const len = bytes.length;
  
  if (
    offset + 8 > len ||
    bytes[offset] !== 0x89 ||
    bytes[offset + 1] !== 0x50 ||
    bytes[offset + 2] !== 0x4E ||
    bytes[offset + 3] !== 0x47 ||
    bytes[offset + 4] !== 0x0D ||
    bytes[offset + 5] !== 0x0A ||
    bytes[offset + 6] !== 0x1A ||
    bytes[offset + 7] !== 0x0A
  ) {
    return -1;
  }
  
  offset += 8;
  
  while (offset + 12 <= len) {
    // 讀取區段長度 (4 bytes, big-endian)
    const chunkLen = ((bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3]) >>> 0;
    
    // 讀取區段類型 (4 bytes)
    const type = String.fromCharCode(bytes[offset + 4], bytes[offset + 5], bytes[offset + 6], bytes[offset + 7]);
    
    // 驗證類型是否皆為合法英文字母，若有任何亂碼、控制字元，表示已跳出合法結構，直接終止
    for (let j = 0; j < 4; j++) {
      const charCode = bytes[offset + 4 + j];
      if (charCode < 65 || (charCode > 90 && charCode < 97) || charCode > 122) {
        return -1;
      }
    }
    
    if (chunkLen < 0 || offset + 12 + chunkLen > len) {
      return -1;
    }
    
    const nextOffset = offset + 12 + chunkLen;
    
    if (type === 'IEND') {
      return nextOffset;
    }
    
    offset = nextOffset;
  }
  
  return -1;
}

/**
 * 取得 OLE CFB 容器中所有可能的二進位流（相容於 .doc 與 .xls），
 * 以確保解碼過程中不漏掉任何分頁 (Sheet)、文字、欄位或內嵌 OLE 物件中的圖片
 */
function getLegacyStreamsAndPayloads(arrayBuffer: ArrayBuffer, fileName: string): { name: string; bytes: Uint8Array }[] {
  const payloads: { name: string; bytes: Uint8Array }[] = [];
  const isXls = fileName.toLowerCase().endsWith('.xls');
  
  let cfbParsed = false;
  try {
    const cfbObj = (CFB as any).default || CFB;
    const container = cfbObj.read(new Uint8Array(arrayBuffer), { type: 'array' });
    const fileIndex = container.FileIndex || [];
    
    // 用於記錄是否成功產生了重組後的繪圖串流
    let reassembledAdded = false;
    
    // 第一階段：先尋找並重組 Workbook 繪圖串流 (針對 XLS)
    if (isXls) {
      for (const entry of fileIndex) {
        if (entry && entry.content && entry.content.length > 0) {
          const name = entry.name || '';
          if (name === 'Workbook' || name === 'Book' || name === 'WORKBOOK' || name === 'BOOK') {
            const bytes = new Uint8Array(entry.content);
            const reassembled = reassembleXlsDrawingStream(bytes);
            if (reassembled && reassembled.length > 0) {
              payloads.push({ name: `Stream_${name}_Reassembled_Drawing`, bytes: reassembled });
              reassembledAdded = true;
            }
          }
        }
      }
    }
    
    // 第二階段：將其他所有非重複且非原始 Workbook/Book 的流加入
    for (const entry of fileIndex) {
      if (entry && entry.content && entry.content.length > 0) {
        const name = entry.name || '';
        const bytes = new Uint8Array(entry.content);
        cfbParsed = true;
        
        const isWorkbookName = name === 'Workbook' || name === 'Book' || name === 'WORKBOOK' || name === 'BOOK';
        
        // 如果已經添加了重組後的 Drawing 串流，就絕對不把原始有 CONTINUE 紀錄干擾的 Workbook 串流加入
        if (isXls && isWorkbookName && reassembledAdded) {
          continue;
        }
        
        payloads.push({ name: `Stream_${name}`, bytes });
      }
    }
  } catch (e) {
    console.error('解析 OLE CFB 容器失敗，將僅使用原始檔案進行雕刻:', e);
  }
  
  // 3. 兜底策略：只有在 OLE 容器解析失敗，或解析出來沒有任何有效 Payload 時，才使用 RawFile 進行雕刻。
  // 這能 100% 避免當能正常解析 OLE 時，因 RawFile 中有 OLE 扇區指針或 sector 碎片導致的「花屏/損壞/重複」圖片！
  if (!cfbParsed || payloads.length === 0) {
    payloads.push({ name: 'RawFile', bytes: new Uint8Array(arrayBuffer) });
  }
  
  return payloads;
}

/**
 * 產生圖片內容特徵指紋，用於進行記憶體級的高效不重複去重
 */
function getImageFingerprint(bytes: Uint8Array): string {
  const len = bytes.length;
  if (len < 200) {
    return `${len}-${bytes.join(',')}`;
  }
  // 使用檔案大小、頭部 100 bytes 及尾部 100 bytes 產生唯一的快速雜湊
  const first100 = bytes.subarray(0, 100);
  const last100 = bytes.subarray(len - 100, len);
  let hash1 = 0;
  for (let i = 0; i < 100; i++) {
    hash1 = (hash1 * 31 + first100[i]) | 0;
  }
  let hash2 = 0;
  for (let i = 0; i < 100; i++) {
    hash2 = (hash2 * 31 + last100[i]) | 0;
  }
  return `${len}-${hash1}-${hash2}`;
}

/**
 * 嘗試從指定 start 索引開始雕刻 JPEG 圖片，優先採用結構特徵解析器，100% 精準取得圖片大小以安全跳過整個檔案，防止被壓縮數據中的假 SOI/EOI 訊號干擾
 */
async function tryCarveJpeg(
  bytes: Uint8Array,
  start: number,
  seenFingerprints: Set<string>,
  targetDirHandle: FileSystemDirectoryHandle,
  baseName: string,
  count: number,
  sheetNames: string[] = [],
  xlsBses: parsedBse[] = [],
  carvedIdx: number = 0
): Promise<{ success: boolean; saved: boolean; end: number; savedCount?: number }> {
  const len = bytes.length;
  const parsedEnd = findJpegEnd(bytes, start);

  // 1. 優先使用精準的結構解析器
  if (parsedEnd !== -1 && parsedEnd <= len) {
    const imageBytes = bytes.subarray(start, parsedEnd);
    const fingerprint = getImageFingerprint(imageBytes);

    const blob = new Blob([imageBytes], { type: 'image/jpeg' });
    const isValid = await isValidBrowserImage(blob);
    if (isValid) {
      seenFingerprints.add(fingerprint);
      
      let cRef = 1;
      if (xlsBses && xlsBses.length > 0) {
        cRef = getXlsImageRefCount(imageBytes.length, xlsBses, carvedIdx);
      }

      for (let r = 0; r < cRef; r++) {
        const prefix = String(count + 1 + r).padStart(3, '0');
        if (sheetNames.length > 0) {
          for (const sheet of sheetNames) {
            const sanitizedSheet = sanitizeFilename(sheet);
            const fileName = `${baseName}_${sanitizedSheet}_image${prefix}.jpg`;
            await safeWriteFile(targetDirHandle, fileName, blob);
          }
        } else {
          const fileName = `${baseName}_image${prefix}.jpg`;
          await safeWriteFile(targetDirHandle, fileName, blob);
        }
      }

      const totalSaved = cRef * (sheetNames.length > 0 ? sheetNames.length : 1);
      return { success: true, saved: true, end: parsedEnd, savedCount: totalSaved };
    }
  }

  // 2. 只有結構解析失敗時，才使用萬一的備用「搜尋 FF D9」容錯
  const potentialEnds: number[] = [];
  let offset = start + 2;
  const maxSearch = Math.min(len, start + 12 * 1024 * 1024); // 最大 12MB
  while (offset < maxSearch - 1) {
    const idx = bytes.indexOf(0xFF, offset);
    if (idx === -1 || idx >= maxSearch - 1) {
      break;
    }
    if (bytes[idx + 1] === 0xD9) {
      const endCandidate = idx + 2;
      if (!potentialEnds.includes(endCandidate)) {
        potentialEnds.push(endCandidate);
      }
    }
    offset = idx + 1;
  }

  potentialEnds.sort((a, b) => a - b);

  for (const end of potentialEnds) {
    const imageBytes = bytes.subarray(start, end);
    const fingerprint = getImageFingerprint(imageBytes);

    const blob = new Blob([imageBytes], { type: 'image/jpeg' });
    const isValid = await isValidBrowserImage(blob);
    if (isValid) {
      seenFingerprints.add(fingerprint);
      
      let cRef = 1;
      if (xlsBses && xlsBses.length > 0) {
        cRef = getXlsImageRefCount(imageBytes.length, xlsBses, carvedIdx);
      }

      for (let r = 0; r < cRef; r++) {
        const prefix = String(count + 1 + r).padStart(3, '0');
        if (sheetNames.length > 0) {
          for (const sheet of sheetNames) {
            const sanitizedSheet = sanitizeFilename(sheet);
            const fileName = `${baseName}_${sanitizedSheet}_image${prefix}.jpg`;
            await safeWriteFile(targetDirHandle, fileName, blob);
          }
        } else {
          const fileName = `${baseName}_image${prefix}.jpg`;
          await safeWriteFile(targetDirHandle, fileName, blob);
        }
      }

      const totalSaved = cRef * (sheetNames.length > 0 ? sheetNames.length : 1);
      return { success: true, saved: true, end, savedCount: totalSaved };
    }
  }

  return { success: false, saved: false, end: -1 };
}

/**
 * 嘗試從指定 start 索引開始雕刻 PNG 圖片，優先採用結構特徵解析器，精準跳過整個檔案
 */
async function tryCarvePng(
  bytes: Uint8Array,
  start: number,
  seenFingerprints: Set<string>,
  targetDirHandle: FileSystemDirectoryHandle,
  baseName: string,
  count: number,
  sheetNames: string[] = [],
  xlsBses: parsedBse[] = [],
  carvedIdx: number = 0
): Promise<{ success: boolean; saved: boolean; end: number; savedCount?: number }> {
  const len = bytes.length;
  const parsedEnd = findPngEnd(bytes, start);

  // 1. 優先使用精準的結構解析器
  if (parsedEnd !== -1 && parsedEnd <= len) {
    const imageBytes = bytes.subarray(start, parsedEnd);
    const fingerprint = getImageFingerprint(imageBytes);

    const blob = new Blob([imageBytes], { type: 'image/png' });
    const isValid = await isValidBrowserImage(blob);
    if (isValid) {
      seenFingerprints.add(fingerprint);
      
      let cRef = 1;
      if (xlsBses && xlsBses.length > 0) {
        cRef = getXlsImageRefCount(imageBytes.length, xlsBses, carvedIdx);
      }

      for (let r = 0; r < cRef; r++) {
        const prefix = String(count + 1 + r).padStart(3, '0');
        if (sheetNames.length > 0) {
          for (const sheet of sheetNames) {
            const sanitizedSheet = sanitizeFilename(sheet);
            const fileName = `${baseName}_${sanitizedSheet}_image${prefix}.png`;
            await safeWriteFile(targetDirHandle, fileName, blob);
          }
        } else {
          const fileName = `${baseName}_image${prefix}.png`;
          await safeWriteFile(targetDirHandle, fileName, blob);
        }
      }

      const totalSaved = cRef * (sheetNames.length > 0 ? sheetNames.length : 1);
      return { success: true, saved: true, end: parsedEnd, savedCount: totalSaved };
    }
  }

  // 2. 只有結構解析失敗時，才使用萬一的備用「搜尋 IEND」容錯
  const potentialEnds: number[] = [];
  let offset = start + 8;
  const maxSearch = Math.min(len, start + 15 * 1024 * 1024); // 最大 15MB
  while (offset < maxSearch - 3) {
    const idx = bytes.indexOf(0x49, offset); // 'I'
    if (idx === -1 || idx >= maxSearch - 3) {
      break;
    }
    if (
      bytes[idx + 1] === 0x45 && // 'E'
      bytes[idx + 2] === 0x4E && // 'N'
      bytes[idx + 3] === 0x44    // 'D'
    ) {
      const endCandidate = idx + 8; // IEND (4) + CRC (4)
      if (endCandidate <= len && !potentialEnds.includes(endCandidate)) {
        potentialEnds.push(endCandidate);
      }
    }
    offset = idx + 1;
  }

  potentialEnds.sort((a, b) => a - b);

  for (const end of potentialEnds) {
    const imageBytes = bytes.subarray(start, end);
    const fingerprint = getImageFingerprint(imageBytes);

    const blob = new Blob([imageBytes], { type: 'image/png' });
    const isValid = await isValidBrowserImage(blob);
    if (isValid) {
      seenFingerprints.add(fingerprint);
      
      let cRef = 1;
      if (xlsBses && xlsBses.length > 0) {
        cRef = getXlsImageRefCount(imageBytes.length, xlsBses, carvedIdx);
      }

      for (let r = 0; r < cRef; r++) {
        const prefix = String(count + 1 + r).padStart(3, '0');
        if (sheetNames.length > 0) {
          for (const sheet of sheetNames) {
            const sanitizedSheet = sanitizeFilename(sheet);
            const fileName = `${baseName}_${sanitizedSheet}_image${prefix}.png`;
            await safeWriteFile(targetDirHandle, fileName, blob);
          }
        } else {
          const fileName = `${baseName}_image${prefix}.png`;
          await safeWriteFile(targetDirHandle, fileName, blob);
        }
      }

      const totalSaved = cRef * (sheetNames.length > 0 ? sheetNames.length : 1);
      return { success: true, saved: true, end, savedCount: totalSaved };
    }
  }

  return { success: false, saved: false, end: -1 };
}

/**
 * 嘗試從指定 start 索引開始雕刻 BMP/DIB 圖片（舊版 Office 常見儲存格式）
 */
async function tryCarveBmp(
  bytes: Uint8Array,
  start: number,
  seenFingerprints: Set<string>,
  targetDirHandle: FileSystemDirectoryHandle,
  baseName: string,
  count: number,
  sheetNames: string[] = [],
  xlsBses: parsedBse[] = [],
  carvedIdx: number = 0
): Promise<{ success: boolean; saved: boolean; end: number; savedCount?: number }> {
  const len = bytes.length;
  if (start + 18 >= len) return { success: false, saved: false, end: -1 };

  const fileSize = bytes[start + 2] | (bytes[start + 3] << 8) | (bytes[start + 4] << 16) | (bytes[start + 5] << 24);

  if (fileSize > 100 && fileSize <= 25 * 1024 * 1024 && start + fileSize <= len) {
    const headerSize = bytes[start + 14] | (bytes[start + 15] << 8) | (bytes[start + 16] << 16) | (bytes[start + 17] << 24);

    if ([12, 40, 52, 56, 64, 108, 124].includes(headerSize)) {
      const imageBytes = bytes.subarray(start, start + fileSize);
      const fingerprint = getImageFingerprint(imageBytes);

      const blob = new Blob([imageBytes], { type: 'image/bmp' });
      const isValid = await isValidBrowserImage(blob);
      if (isValid) {
        seenFingerprints.add(fingerprint);
        
        let cRef = 1;
        if (xlsBses && xlsBses.length > 0) {
          cRef = getXlsImageRefCount(imageBytes.length, xlsBses, carvedIdx);
        }

        for (let r = 0; r < cRef; r++) {
          const prefix = String(count + 1 + r).padStart(3, '0');
          if (sheetNames.length > 0) {
            for (const sheet of sheetNames) {
              const sanitizedSheet = sanitizeFilename(sheet);
              const fileName = `${baseName}_${sanitizedSheet}_image${prefix}.bmp`;
              await safeWriteFile(targetDirHandle, fileName, blob);
            }
          } else {
            const fileName = `${baseName}_image${prefix}.bmp`;
            await safeWriteFile(targetDirHandle, fileName, blob);
          }
        }

        const totalSaved = cRef * (sheetNames.length > 0 ? sheetNames.length : 1);
        return { success: true, saved: true, end: start + fileSize, savedCount: totalSaved };
      }
    }
  }

  return { success: false, saved: false, end: -1 };
}

interface parsedBse {
  size: number;
  cRef: number;
  rgbUid: string;
}

/**
 * 掃描 XLS 的二進位資料，提取所有 BSE (B-Store Entry) 紀錄與其參照計數 (Duplicate reference counts)
 */
function parseXlsBseRecords(bytes: Uint8Array): parsedBse[] {
  const bses: parsedBse[] = [];
  const len = bytes.length;
  
  for (let i = 2; i < len - 42; i++) {
    // 尋找 Escher 紀錄類型 0xF007 (little-endian 為 0x07, 0xF0)
    if (bytes[i] === 0x07 && bytes[i + 1] === 0xF0) {
      const recLength = bytes[i + 2] | (bytes[i + 3] << 8) | (bytes[i + 4] << 16) | (bytes[i + 5] << 24);
      
      // BSE 紀錄負載長度至少應為 36 bytes (在 OLE 中)，且不應超出剩餘空間
      if (recLength >= 36 && recLength <= len - i - 6) {
        const payloadOffset = i + 6;
        const btWin32 = bytes[payloadOffset];
        const btMacOS = bytes[payloadOffset + 1];
        
        // 驗證圖片類型是否在合理區間 (通常 0-10 為有效類型)
        if (btWin32 <= 12 && btMacOS <= 12) {
          // rgbUid (16 bytes，為圖片資料的 MD5 雜湊)
          const rgbUidBytes = bytes.subarray(payloadOffset + 2, payloadOffset + 18);
          const rgbUid = Array.from(rgbUidBytes).map(b => b.toString(16).padStart(2, '0')).join('');
          
          // size (4 bytes，圖片的二進位大小)
          const size = bytes[payloadOffset + 20] | (bytes[payloadOffset + 21] << 8) | (bytes[payloadOffset + 22] << 16) | (bytes[payloadOffset + 23] << 24);
          
          // cRef (4 bytes，此圖片在工作表中的總參照/出現次數)
          const cRef = bytes[payloadOffset + 24] | (bytes[payloadOffset + 25] << 8) | (bytes[payloadOffset + 26] << 16) | (bytes[payloadOffset + 27] << 24);
          
          if (size > 0 && cRef > 0 && cRef < 100000) {
            bses.push({ size, cRef, rgbUid });
            // 跳過當前 BSE 紀錄本體，加速尋找下一個
            i += 6 + recLength - 1;
          }
        }
      }
    }
  }
  return bses;
}

/**
 * 依據雕刻出的圖片位元組長度與雕刻順序索引，從 BSE 列表匹配出該圖片的正確參照次數
 */
function getXlsImageRefCount(
  carvedLength: number,
  xlsBses: parsedBse[],
  carvedIdx: number
): number {
  if (!xlsBses || xlsBses.length === 0) return 1;
  
  // 1. 優先透過檔案大小匹配 (容許一些雕刻結尾的微小解析差異，設定 128 bytes 容差)
  const matched = xlsBses.find(b => Math.abs(b.size - carvedLength) < 128);
  if (matched) {
    return Math.max(1, matched.cRef);
  }
  
  // 2. 兜底：依據順序索引直接匹配
  if (carvedIdx < xlsBses.length) {
    return Math.max(1, xlsBses[carvedIdx].cRef);
  }
  
  return 1;
}

/**
 * 從舊版二進位格式 (.doc, .xls) 中透過二進位特徵 (Magic Numbers/Carving) 擷取 JPEG, PNG 與 BMP 圖片
 */
export async function extractImagesFromLegacyOffice(file: File, targetDirHandle: FileSystemDirectoryHandle): Promise<number> {
  const arrayBuffer = await file.arrayBuffer();

  let count = 0;
  const baseName = getFileBaseName(file.name);
  const seenFingerprints = new Set<string>();

  const isXls = file.name.toLowerCase().endsWith('.xls');
  let sheetNames: string[] = [];
  let xlsBses: parsedBse[] = [];

  if (isXls) {
    try {
      sheetNames = parseXlsSheetNames(new Uint8Array(arrayBuffer));
    } catch (e) {
      console.warn('解析 XLS 工作表名稱失敗:', e);
    }
    try {
      xlsBses = parseXlsBseRecords(new Uint8Array(arrayBuffer));
      console.log('解析出 XLS BSE 紀錄數量:', xlsBses.length, xlsBses);
    } catch (e) {
      console.warn('解析 XLS BSE 紀錄失敗:', e);
    }
  }

  // 取得所有待雕刻的二進位來源
  const sources = getLegacyStreamsAndPayloads(arrayBuffer, file.name);

  // 優先將重組後的 Drawing 串流排在最前面，確保以此連續串流獲取最佳解碼畫質的圖片
  sources.sort((a, b) => {
    const aIsReassembled = a.name.endsWith('_Reassembled_Drawing');
    const bIsReassembled = b.name.endsWith('_Reassembled_Drawing');
    if (aIsReassembled && !bIsReassembled) return -1;
    if (!aIsReassembled && bIsReassembled) return 1;
    return 0;
  });

  let carvedIdx = 0;

  // 對每個來源進行 JPEG, PNG 與 BMP 雕刻
  for (const source of sources) {
    const bytes = source.bytes;
    const len = bytes.length;
    let i = 0;

    while (i < len - 4) {
      // 1. JPEG 雕刻
      if (bytes[i] === 0xFF && bytes[i + 1] === 0xD8) {
        const res = await tryCarveJpeg(bytes, i, seenFingerprints, targetDirHandle, baseName, count, sheetNames, xlsBses, carvedIdx);
        if (res.success) {
          if (res.saved) {
            count += res.savedCount || 1;
            carvedIdx++;
          }
          i = res.end;
          continue;
        }
      }

      // 2. PNG 雕刻
      if (
        i < len - 8 &&
        bytes[i] === 0x89 &&
        bytes[i + 1] === 0x50 &&
        bytes[i + 2] === 0x4E &&
        bytes[i + 3] === 0x47 &&
        bytes[i + 4] === 0x0D &&
        bytes[i + 5] === 0x0A &&
        bytes[i + 6] === 0x1A &&
        bytes[i + 7] === 0x0A
      ) {
        const res = await tryCarvePng(bytes, i, seenFingerprints, targetDirHandle, baseName, count, sheetNames, xlsBses, carvedIdx);
        if (res.success) {
          if (res.saved) {
            count += res.savedCount || 1;
            carvedIdx++;
          }
          i = res.end;
          continue;
        }
      }

      // 3. BMP 雕刻
      if (bytes[i] === 0x42 && bytes[i + 1] === 0x4D) {
        const res = await tryCarveBmp(bytes, i, seenFingerprints, targetDirHandle, baseName, count, sheetNames, xlsBses, carvedIdx);
        if (res.success) {
          if (res.saved) {
            count += res.savedCount || 1;
            carvedIdx++;
          }
          i = res.end;
          continue;
        }
      }

      i++;
    }
  }

  return count;
}
