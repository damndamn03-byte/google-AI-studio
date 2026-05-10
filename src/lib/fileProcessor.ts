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
  const zip = await JSZip.loadAsync(arrayBuffer);
  
  let count = 0;
  const isDocx = file.name.endsWith('.docx');
  const mediaPath = isDocx ? 'word/media/' : 'xl/media/';
  
  const mediaFiles = Object.keys(zip.files).filter(path => path.startsWith(mediaPath));
  
  for (const path of mediaFiles) {
    const zipFile = zip.files[path];
    if (zipFile.dir) continue;
    
    const blob = await zipFile.async('blob');
    const fileName = path.split('/').pop() || `image_${count}`;
    const fileHandle = await targetDirHandle.getFileHandle(fileName, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    count++;
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
