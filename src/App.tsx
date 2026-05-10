/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { 
  FolderOpen, 
  FileImage, 
  CheckCircle2, 
  AlertCircle, 
  ChevronRight, 
  Download, 
  Loader2, 
  FileText,
  FileSpreadsheet,
  Image as ImageIcon,
  FolderSync,
  ExternalLink,
  Square,
  Bell,
  BellOff
} from 'lucide-react';
import { motion } from 'motion/react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Progress } from '@/components/ui/progress';
import { Badge } from '@/components/ui/badge';
import { ScrollArea } from '@/components/ui/scroll-area';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { 
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
  DialogFooter,
} from "@/components/ui/dialog";
import { extractImagesFromOffice, extractImagesFromPdf, ProcessingResult } from '@/src/lib/fileProcessor';

export default function App() {
  const [isProcessing, setIsProcessing] = useState(false);
  const [stopRequested, setStopRequested] = useState(false);
  const [results, setResults] = useState<ProcessingResult[]>([]);
  const [progress, setProgress] = useState(0);
  const [statusText, setStatusText] = useState('');
  const [error, setError] = useState<{ message: string, isIframeError?: boolean } | null>(null);
  const [showCompleteDialog, setShowCompleteDialog] = useState(false);
  const [notificationsEnabled, setNotificationsEnabled] = useState(false);

  const isInIframe = window.self !== window.top;

  // Initialize notification state
  useEffect(() => {
    if ('Notification' in window) {
      setNotificationsEnabled(Notification.permission === 'granted');
    }
  }, []);

  const requestNotificationPermission = async () => {
    if ('Notification' in window) {
      const permission = await Notification.requestPermission();
      setNotificationsEnabled(permission === 'granted');
    }
  };

  const playCompletionSound = () => {
    try {
      const audioCtx = new (window.AudioContext || (window as any).webkitAudioContext)();
      const oscillator = audioCtx.createOscillator();
      const gainNode = audioCtx.createGain();

      oscillator.connect(gainNode);
      gainNode.connect(audioCtx.destination);

      oscillator.type = 'sine';
      oscillator.frequency.setValueAtTime(880, audioCtx.currentTime); // A5
      oscillator.frequency.exponentialRampToValueAtTime(1320, audioCtx.currentTime + 0.1); // E6

      gainNode.gain.setValueAtTime(0.1, audioCtx.currentTime);
      gainNode.gain.exponentialRampToValueAtTime(0.01, audioCtx.currentTime + 0.3);

      oscillator.start();
      oscillator.stop(audioCtx.currentTime + 0.3);
    } catch (e) {
      console.warn('Audio playback failed', e);
    }
  };

  const notifyComplete = (stats: any) => {
    playCompletionSound();
    setShowCompleteDialog(true);

    if (notificationsEnabled && Notification.permission === 'granted') {
      new Notification('Image Extractor Pro 任務完成', {
        body: `成功處理 ${stats.total} 個檔案，共擷取 ${results.reduce((acc, r) => acc + r.imagesExtracted, 0)} 張圖片。`,
        icon: 'https://cdn-icons-png.flaticon.com/512/3342/3342137.png'
      });
    }
  };

  const handleStop = () => {
    if (isProcessing) {
      setStopRequested(true);
      setStatusText('正等待當前檔案處理完畢後停止...');
    }
  };

  const processDirectory = async () => {
    try {
      if (!('showDirectoryPicker' in window)) {
        throw new Error('您的瀏覽器不支援 File System Access API。請使用 Chrome、Edge 或 Opera 並確保在安全連線 (HTTPS) 下使用。');
      }

      const dirHandle = await (window as any).showDirectoryPicker({
        mode: 'readwrite'
      }).catch((e: Error) => {
        if (e.name === 'SecurityError' || e.message.includes('sub frames')) {
          const err = new Error('安全性限制：無法在預覽視窗中直接讀取資料夾。請點擊右上角或下方的「在新分頁開啟」按鈕。');
          (err as any).isIframeError = true;
          throw err;
        }
        throw e;
      });

      setIsProcessing(true);
      setStopRequested(false);
      setResults([]);
      setProgress(0);
      setError(null);

      const filesToProcess: { handle: FileSystemFileHandle, path: string }[] = [];
      
      const scanDir = async (handle: FileSystemDirectoryHandle, currentPath: string = '') => {
        for await (const entry of (handle as any).values()) {
          const entryPath = currentPath ? `${currentPath}/${entry.name}` : entry.name;
          if (entry.kind === 'file') {
            const ext = entry.name.toLowerCase().split('.').pop();
            if (['docx', 'xlsx', 'pdf', 'doc', 'xls'].includes(ext || '')) {
              filesToProcess.push({ handle: entry as FileSystemFileHandle, path: entryPath });
            }
          } else if (entry.kind === 'directory') {
            await scanDir(entry as FileSystemDirectoryHandle, entryPath);
          }
        }
      };

      setStatusText('正在掃描資料夾...');
      await scanDir(dirHandle);

      if (filesToProcess.length === 0) {
        setStatusText('未找到支援的檔案類型。');
        setIsProcessing(false);
        return;
      }

      const total = filesToProcess.length;
      const newResults: ProcessingResult[] = [];

      for (let i = 0; i < total; i++) {
        // Check if user requested to stop
        if (stopRequested) {
          setStatusText(`已停止作業。共處理了 ${newResults.length} 個檔案。`);
          break;
        }

        const { handle, path: relativePath } = filesToProcess[i];
        const file = await handle.getFile();
        const ext = file.name.toLowerCase().split('.').pop();
        
        setStatusText(`正在處理 (${i + 1}/${total}): ${file.name}`);
        setProgress(Math.round(((i + 1) / total) * 100));

        // Get the parent directory handle
        let currentDir = dirHandle;
        const pathParts = relativePath.split('/');
        pathParts.pop(); // Remove filename
        
        for (const part of pathParts) {
          try {
            currentDir = await currentDir.getDirectoryHandle(part);
          } catch(e) {
            console.error(`Failed to get directory handle for ${part}`, e);
          }
        }

        try {
          let count = 0;
          let status: ProcessingResult['status'] = 'success';
          let message = '';

          if (['doc', 'xls'].includes(ext || '')) {
            status = 'legacy_skipped';
            message = '舊式二進位格式 (.doc, .xls) 目前不支援自動擷取，請轉檔為 .docx 或 .xlsx 再試。';
          } else {
            // Create target directory named after the file
            const targetDirName = `${file.name}_extracted_images`;
            const targetDirHandle = await currentDir.getDirectoryHandle(targetDirName, { create: true });

            if (ext === 'docx' || ext === 'xlsx') {
              count = await extractImagesFromOffice(file, targetDirHandle);
            } else if (ext === 'pdf') {
              const { count: pdfCount, isScanned } = await extractImagesFromPdf(file, targetDirHandle);
              count = pdfCount;
              if (isScanned) {
                status = 'skipped_scan';
                message = 'PDF 為掃描型式（無文字資訊），已跳過。';
                // Remove the empty created directory
                try {
                   await currentDir.removeEntry(targetDirName, { recursive: true });
                } catch(e) {}
              }
            }
          }

          const result: ProcessingResult = {
            fileName: file.name,
            filePath: relativePath,
            status,
            imagesExtracted: count,
            message
          };
          newResults.push(result);
          setResults([...newResults]);
        } catch (err) {
          newResults.push({
            fileName: file.name,
            filePath: relativePath,
            status: 'error',
            imagesExtracted: 0,
            message: err instanceof Error ? err.message : '未知錯誤'
          });
          setResults([...newResults]);
        }

        // 稍微延遲以讓瀏覽器 UI 線程有喘息空間，減少大量處理時崩潰機率
        await new Promise(resolve => setTimeout(resolve, 50));
      }

      const currentStats = {
        total: newResults.length,
        success: newResults.filter(r => r.status === 'success' && r.imagesExtracted > 0).length,
      };

      setStatusText(stopRequested ? '作業已由使用者中斷。' : '處理完成！');
      setIsProcessing(false);
      setStopRequested(false);
      notifyComplete(currentStats);
    } catch (err) {
      console.error(err);
      setError({
        message: err instanceof Error ? err.message : '發生錯誤',
        isIframeError: (err as any).isIframeError
      });
      setIsProcessing(false);
      setStopRequested(false);
    }
  };

  const downloadReport = () => {
    const csvContent = "data:text/csv;charset=utf-8,\uFEFF" 
      + "檔案名稱,路徑,狀態,擷取圖片數量,備註\n"
      + results.map(r => `"${r.fileName}","${r.filePath}","${r.status}",${r.imagesExtracted},"${r.message || ''}"`).join("\n");
    
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "處理結果報表.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const stats = {
    total: results.length,
    success: results.filter(r => r.status === 'success' && r.imagesExtracted > 0).length,
    scanned: results.filter(r => r.status === 'skipped_scan').length,
    noImages: results.filter(r => r.status === 'success' && r.imagesExtracted === 0).length,
    legacy: results.filter(r => r.status === 'legacy_skipped').length,
    error: results.filter(r => r.status === 'error').length,
  };

  return (
    <div className="min-h-screen bg-neutral-50 p-4 md:p-8 font-sans">
      <div className="max-w-5xl mx-auto space-y-6">
        {/* Header */}
        <header className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          <div>
            <h1 className="text-3xl font-bold tracking-tight text-neutral-900">Image Extractor Pro</h1>
            <p className="text-neutral-500 font-medium mt-1">文件圖片自動擷取工具</p>
          </div>
          <div className="flex gap-2">
            {!isInIframe && 'Notification' in window && (
              <Button
                variant="outline"
                size="icon"
                onClick={requestNotificationPermission}
                className={`h-12 w-12 rounded-xl transition-all ${notificationsEnabled ? 'text-green-600 bg-green-50' : 'text-neutral-400'}`}
                title={notificationsEnabled ? "桌面通知已開啟" : "開啟桌面通知"}
              >
                {notificationsEnabled ? <Bell className="h-5 w-5" /> : <BellOff className="h-5 w-5" />}
              </Button>
            )}
            {isInIframe && (
              <Button 
                variant="outline"
                className="h-12 px-6 rounded-xl border-neutral-200"
                onClick={() => window.open(window.location.href, '_blank')}
              >
                <ExternalLink className="mr-2 h-5 w-5" />
                在新分頁開啟
              </Button>
            )}
            <Button 
              onClick={processDirectory} 
              disabled={isProcessing}
              className="bg-neutral-900 text-white hover:bg-neutral-800 h-12 px-6 rounded-xl transition-all shadow-sm active:scale-95"
            >
              {isProcessing ? <Loader2 className="mr-2 h-5 w-5 animate-spin" /> : <FolderSync className="mr-2 h-5 w-5" />}
              {isProcessing ? '正在處理中...' : '選擇資料夾開始處理'}
            </Button>
            {isProcessing && (
              <Button 
                variant="destructive" 
                onClick={handleStop}
                disabled={stopRequested}
                className="h-12 px-6 rounded-xl bg-red-600 hover:bg-red-700 text-white shadow-sm"
              >
                <Square className="mr-2 h-5 w-5 fill-current" />
                停止
              </Button>
            )}
            {results.length > 0 && !isProcessing && (
              <Button variant="outline" onClick={downloadReport} className="h-12 px-6 rounded-xl border-neutral-200">
                <Download className="mr-2 h-5 w-5" />
                下載報表
              </Button>
            )}
          </div>
        </header>

        {isInIframe && !error && (
          <Alert className="bg-amber-50 border-amber-200 text-amber-800">
            <AlertCircle className="h-4 w-4 text-amber-600" />
            <AlertTitle className="font-bold">提醒</AlertTitle>
            <AlertDescription>
              由於瀏覽器安全政策，請點擊右上方「在新分頁開啟」以正常使用資料夾存取功能。
            </AlertDescription>
          </Alert>
        )}

        {error && (
          <Alert variant="destructive" className="bg-red-50 border-red-200">
            <AlertCircle className="h-4 w-4" />
            <AlertTitle>錯誤</AlertTitle>
            <AlertDescription className="space-y-4">
              <p>{error.message}</p>
              {error.isIframeError && (
                <Button 
                  size="sm" 
                  variant="outline" 
                  className="mt-2 bg-white text-red-600 border-red-200 hover:bg-red-50"
                  onClick={() => window.open(window.location.href, '_blank')}
                >
                  <ExternalLink className="mr-2 h-4 w-4" />
                  立即在新分頁開啟
                </Button>
              )}
            </AlertDescription>
          </Alert>
        )}

        {/* Stats Cards */}
        {results.length > 0 && (
          <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
            <StatCard label="總檔案" value={stats.total} icon={<FileText className="w-4 h-4" />} />
            <StatCard label="擷取成功" value={stats.success} color="text-green-600" icon={<CheckCircle2 className="w-4 h-4" />} />
            <StatCard label="掃描型 PDF" value={stats.scanned} color="text-amber-600" icon={<ImageIcon className="w-4 h-4" />} />
            <StatCard label="舊式格式" value={stats.legacy} color="text-blue-600" icon={<AlertCircle className="w-4 h-4" />} />
            <StatCard label="錯誤" value={stats.error} color="text-red-600" icon={<AlertCircle className="w-4 h-4" />} />
          </div>
        )}

        {/* Progress */}
        {isProcessing && (
          <Card className="border-none shadow-sm ring-1 ring-neutral-200">
            <CardHeader className="pb-2">
              <div className="flex justify-between items-end">
                <CardTitle className="text-sm font-medium">{statusText}</CardTitle>
                <span className="text-xs text-neutral-500 font-mono font-bold">{progress}%</span>
              </div>
            </CardHeader>
            <CardContent>
              <Progress value={progress} className="h-2 bg-neutral-100" />
            </CardContent>
          </Card>
        )}

        {/* Results List */}
        <Card className="border-none shadow-md overflow-hidden bg-white ring-1 ring-neutral-200">
          <CardHeader className="bg-neutral-50/50 border-b border-neutral-100">
            <CardTitle className="text-lg">處理進度記錄</CardTitle>
            <CardDescription>
              處理後的圖片將存放於原檔案所在位置的同名資料夾中。
            </CardDescription>
          </CardHeader>
          <CardContent className="p-0">
            <ScrollArea className="h-[500px]">
              {results.length === 0 ? (
                <div className="flex flex-col items-center justify-center h-[400px] text-neutral-400 space-y-4">
                  <div className="bg-neutral-100 p-6 rounded-full">
                    <FolderOpen className="w-12 h-12 opacity-40" />
                  </div>
                  <div className="text-center">
                    <p className="font-semibold text-neutral-600">尚未選擇資料夾</p>
                    <p className="text-sm mt-1">點擊上方按鈕選擇包含 Office 或 PDF 檔案的資料夾</p>
                  </div>
                </div>
              ) : (
                <div className="divide-y divide-neutral-100">
                  {results.map((result, idx) => (
                    <motion.div 
                      key={idx}
                      initial={{ opacity: 0, x: -10 }}
                      animate={{ opacity: 1, x: 0 }}
                      className="p-4 hover:bg-neutral-50 transition-colors flex items-center justify-between gap-4"
                    >
                      <div className="flex items-center gap-3 min-w-0">
                        <FileIcon extension={result.fileName.split('.').pop() || ''} />
                        <div className="truncate">
                          <p className="font-medium text-sm text-neutral-900 truncate">
                            {result.fileName}
                          </p>
                          <p className="text-xs text-neutral-400 truncate font-mono">
                            {result.filePath}
                          </p>
                        </div>
                      </div>
                      <div className="flex items-center gap-3 shrink-0">
                        {result.status === 'success' && (
                          <div className="flex items-center gap-2">
                            <span className="text-sm text-green-600 font-bold">{result.imagesExtracted} 張</span>
                            <Badge className="bg-green-100 text-green-700 hover:bg-green-100 border-none px-2 rounded-md">完成</Badge>
                          </div>
                        )}
                        {result.status === 'skipped_scan' && (
                          <Badge variant="outline" className="text-amber-600 border-amber-200 bg-amber-50 rounded-md">掃描型 PDF</Badge>
                        )}
                        {result.status === 'legacy_skipped' && (
                          <Badge variant="outline" className="text-blue-600 border-blue-200 bg-blue-50 rounded-md">舊式格式</Badge>
                        )}
                        {result.status === 'error' && (
                          <Badge variant="destructive" className="bg-red-50 text-red-600 hover:bg-red-50 border-none rounded-md">失敗</Badge>
                        )}
                        <ChevronRight className="w-4 h-4 text-neutral-200" />
                      </div>
                    </motion.div>
                  ))}
                </div>
              )}
            </ScrollArea>
          </CardContent>
        </Card>

        {/* Info */}
        <footer className="text-center text-xs text-neutral-400 py-8 border-t border-neutral-100">
          <p>支援格式: .docx, .xlsx, .pdf, .doc (標記), .xls (標記)</p>
          <p className="mt-1 opacity-70">
            注意: 基於隱私與安全，本工具完全在您的瀏覽器端運行，檔案不會傳送至雲端。
          </p>
        </footer>
      </div>

      {/* Completion Dialog */}
      <Dialog open={showCompleteDialog} onOpenChange={setShowCompleteDialog}>
        <DialogContent className="sm:max-w-md rounded-2xl">
          <DialogHeader>
            <div className="mx-auto bg-green-100 p-3 rounded-full w-fit mb-4">
              <CheckCircle2 className="h-8 w-8 text-green-600" />
            </div>
            <DialogTitle className="text-center text-xl">任務處理完成！</DialogTitle>
            <DialogDescription className="text-center">
              所有文件已按照您的要求完成圖片擷取作業。
            </DialogDescription>
          </DialogHeader>
          <div className="grid grid-cols-2 gap-4 py-4">
            <div className="bg-neutral-50 p-4 rounded-xl text-center">
              <p className="text-xs text-neutral-400 font-bold uppercase tracking-wider mb-1">處理總數</p>
              <p className="text-2xl font-bold">{stats.total}</p>
            </div>
            <div className="bg-neutral-50 p-4 rounded-xl text-center">
              <p className="text-xs text-neutral-400 font-bold uppercase tracking-wider mb-1">成功擷取</p>
              <p className="text-2xl font-bold text-green-600">{stats.success}</p>
            </div>
          </div>
          <DialogFooter className="sm:justify-center">
            <Button 
              type="button" 
              className="w-full bg-neutral-900 text-white rounded-xl h-11"
              onClick={() => setShowCompleteDialog(false)}
            >
              我知道了
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}

function StatCard({ label, value, color = "text-neutral-900", icon }: { label: string, value: number, color?: string, icon: React.ReactNode }) {
  return (
    <Card className="border-none shadow-sm ring-1 ring-neutral-200 bg-white">
      <CardHeader className="p-4 flex flex-row items-center justify-between space-y-0 pb-1">
        <CardTitle className="text-[10px] font-bold text-neutral-400 uppercase tracking-widest">{label}</CardTitle>
        <div className="text-neutral-200">{icon}</div>
      </CardHeader>
      <CardContent className="p-4 pt-0">
        <div className={`text-2xl font-bold tracking-tight ${color}`}>{value}</div>
      </CardContent>
    </Card>
  );
}

function FileIcon({ extension }: { extension: string }) {
  const ext = extension.toLowerCase();
  if (['docx', 'doc'].includes(ext)) return <FileText className="w-8 h-8 text-blue-500 bg-blue-50 p-1.5 rounded-lg" />;
  if (['xlsx', 'xls'].includes(ext)) return <FileSpreadsheet className="w-8 h-8 text-green-600 bg-green-50 p-1.5 rounded-lg" />;
  if (ext === 'pdf') return <FileText className="w-8 h-8 text-red-500 bg-red-50 p-1.5 rounded-lg" />;
  return <FolderOpen className="w-8 h-8 text-neutral-400 bg-neutral-100 p-1.5 rounded-lg" />;
}
