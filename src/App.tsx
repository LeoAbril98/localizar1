import React, { useState, useEffect, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { BrowserMultiFormatReader, Result } from '@zxing/library';
import { 
  Camera, 
  Upload, 
  Search, 
  X, 
  Package, 
  MapPin, 
  Hash, 
  CheckCircle2, 
  AlertCircle,
  RefreshCcw,
  FileSpreadsheet,
  QrCode
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// Utility for tailwind classes
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

interface InventoryItem {
  Código: string | number;
  Modelo: string;
  Local: string;
  Quantidade: string | number;
}

export default function App() {
  const [data, setData] = useState<InventoryItem[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [result, setResult] = useState<InventoryItem | null>(null);
  const [isScanning, setIsScanning] = useState(false);
  const [isCameraReady, setIsCameraReady] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  
  const videoRef = useRef<HTMLVideoElement>(null);
  const codeReaderRef = useRef<BrowserMultiFormatReader | null>(null);

  // Initialize Code Reader
  useEffect(() => {
    codeReaderRef.current = new BrowserMultiFormatReader();
    return () => {
      if (codeReaderRef.current) {
        codeReaderRef.current.reset();
      }
    };
  }, []);

  // Handle Excel Upload
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const jsonData = XLSX.utils.sheet_to_json(ws) as any[];
        
        // Map columns if they use the new format
        const mappedData: InventoryItem[] = jsonData.map(row => ({
          Código: row.Código ?? row.Produto ?? '',
          Modelo: row.Modelo ?? row.Descricao ?? '',
          Local: row.Local ?? '',
          Quantidade: row.Quantidade ?? row.Qtde ?? 0
        }));

        // Basic validation: Check if at least some data was parsed
        if (mappedData.length > 0) {
          const firstRow = mappedData[0];
          // Check if we have at least Código/Produto and Modelo/Descricao
          if (!firstRow.Código && !firstRow.Modelo) {
            setError("Formato de planilha inválido. Certifique-se de ter as colunas: Produto/Código e Descricao/Modelo.");
            setData([]);
          } else {
            setData(mappedData);
            setError(null);
          }
        } else {
          setError("A planilha está vazia.");
        }
      } catch (err) {
        setError("Erro ao ler o arquivo Excel.");
        console.error(err);
      }
    };
    reader.readAsBinaryString(file);
  };

  // Search Logic
  const handleSearch = useCallback((query: string) => {
    if (!data.length) return;
    
    const cleanQuery = query.toString().trim().toLowerCase();
    if (!cleanQuery) {
      setResult(null);
      return;
    }

    const found = data.find(item => 
      item.Código?.toString().toLowerCase() === cleanQuery
    );

    if (found) {
      setResult(found);
      setError(null);
    } else {
      setResult(null);
      setError("Produto não encontrado");
    }
  }, [data]);

  // Barcode Scanning Logic
  const startScanner = useCallback(async () => {
    if (!codeReaderRef.current || !videoRef.current) return;
    setError(null);

    try {
      if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
        throw new Error("Seu navegador não suporta acesso à câmera.");
      }

      const constraints = {
        video: {
          facingMode: 'environment',
          width: { ideal: 1280 },
          height: { ideal: 720 }
        }
      };

      // Explicitly get the stream first to verify it works
      const stream = await navigator.mediaDevices.getUserMedia(constraints);
      
      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        videoRef.current.setAttribute('playsinline', 'true');
        videoRef.current.muted = true;
        // Ensure the video starts playing
        try {
          await videoRef.current.play();
          
          // Wait for video to be actually ready
          if (videoRef.current.readyState < 2) {
            await new Promise((resolve) => {
              videoRef.current!.onloadeddata = resolve;
            });
          }

          // Double check if the stream tracks are active
          const tracks = stream.getVideoTracks();
          if (tracks.length === 0 || tracks[0].readyState !== 'live') {
            throw new Error("A trilha de vídeo não está ativa.");
          }
          
          setIsCameraReady(true);
          await new Promise(resolve => setTimeout(resolve, 300));
        } catch (playErr) {
          console.error("Video play error:", playErr);
        }
      }

      await codeReaderRef.current.decodeFromStream(
        stream,
        videoRef.current,
        (result: Result | null, err?: Error) => {
          if (result) {
            const code = result.getText();
            setSearchQuery(code);
            handleSearch(code);
            stopScanner();
          }
        }
      );
    } catch (err: any) {
      console.error("Camera Error:", err);
      let msg = "Erro ao acessar a câmera.";
      if (err.name === 'NotAllowedError' || err.name === 'PermissionDeniedError') {
        msg = "Acesso negado. Por favor, permita o uso da câmera nas configurações.";
      } else if (err.name === 'NotFoundError') {
        msg = "Câmera não encontrada.";
      }
      setError(msg);
      setIsScanning(false);
    }
  }, [handleSearch]);

  // Effect to start scanner when isScanning becomes true and video element is ready
  useEffect(() => {
    if (isScanning && videoRef.current) {
      startScanner();
    }
  }, [isScanning, startScanner]);

  const stopScanner = () => {
    if (codeReaderRef.current) {
      codeReaderRef.current.reset();
    }
    if (videoRef.current && videoRef.current.srcObject) {
      const stream = videoRef.current.srcObject as MediaStream;
      stream.getTracks().forEach(track => track.stop());
      videoRef.current.srcObject = null;
    }
    setIsScanning(false);
    setIsCameraReady(false);
  };

  const clearSearch = () => {
    setSearchQuery('');
    setResult(null);
    setError(null);
  };

  return (
    <div className="min-h-screen bg-[#F8F9FA] text-[#212529] font-sans selection:bg-blue-100">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-50">
        <div className="max-w-xl mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-blue-600 p-2 rounded-lg shadow-sm">
              <Package className="text-white w-5 h-5" />
            </div>
            <h1 className="font-bold text-lg tracking-tight">StockFinder</h1>
          </div>
          {data.length > 0 && (
            <div className="flex items-center gap-1.5 text-xs font-medium text-emerald-600 bg-emerald-50 px-2 py-1 rounded-full border border-emerald-100">
              <CheckCircle2 className="w-3.5 h-3.5" />
              {data.length} Itens
            </div>
          )}
        </div>
      </header>

      <main className="max-w-xl mx-auto px-4 py-6 space-y-6">
        {/* Step 1: Upload Excel */}
        {data.length === 0 ? (
          <motion.div 
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            className="bg-white rounded-2xl p-8 border border-gray-200 shadow-sm text-center space-y-4"
          >
            <div className="w-16 h-16 bg-blue-50 rounded-full flex items-center justify-center mx-auto">
              <FileSpreadsheet className="text-blue-600 w-8 h-8" />
            </div>
            <div className="space-y-2">
              <h2 className="text-xl font-bold">Carregar Planilha</h2>
              <p className="text-gray-500 text-sm">
                Envie seu arquivo .xlsx ou .csv com as colunas:<br/>
                <span className="font-mono text-[10px] bg-gray-100 px-1.5 py-0.5 rounded block mt-2">
                  Produto (ou Código), Descricao (ou Modelo), Local, Qtde (ou Quantidade)
                </span>
              </p>
            </div>
            <label className="block">
              <span className="sr-only">Escolher arquivo</span>
              <div className="relative group cursor-pointer">
                <input 
                  type="file" 
                  accept=".xlsx, .xls, .csv" 
                  onChange={handleFileUpload}
                  className="hidden"
                />
                <div className="flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3 px-6 rounded-xl transition-all shadow-lg shadow-blue-200 active:scale-95">
                  <Upload className="w-5 h-5" />
                  Selecionar Excel
                </div>
              </div>
            </label>
          </motion.div>
        ) : (
          <div className="space-y-6">
            {/* Search Controls */}
            <div className="space-y-4">
              <div className="flex gap-2">
                <div className="relative flex-1">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 w-5 h-5" />
                  <input
                    type="text"
                    placeholder="Digite o código..."
                    value={searchQuery}
                    onChange={(e) => {
                      setSearchQuery(e.target.value);
                      handleSearch(e.target.value);
                    }}
                    className="w-full pl-10 pr-4 py-3.5 bg-white border border-gray-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition-all shadow-sm"
                  />
                  {searchQuery && (
                    <button 
                      onClick={clearSearch}
                      className="absolute right-3 top-1/2 -translate-y-1/2 p-1 hover:bg-gray-100 rounded-full"
                    >
                      <X className="w-4 h-4 text-gray-400" />
                    </button>
                  )}
                </div>
                <button
                  onClick={() => setIsScanning(true)}
                  className="bg-gray-900 text-white p-4 rounded-xl shadow-lg hover:bg-black transition-all active:scale-95"
                  title="Ler Código de Barras"
                >
                  <QrCode className="w-6 h-6" />
                </button>
              </div>

              <div className="flex items-center justify-between px-1">
                <button 
                  onClick={() => { setData([]); setFileName(null); clearSearch(); }}
                  className="text-xs font-medium text-gray-400 hover:text-red-500 flex items-center gap-1 transition-colors"
                >
                  <RefreshCcw className="w-3 h-3" />
                  Trocar Planilha
                </button>
                {fileName && <span className="text-[10px] text-gray-400 font-mono truncate max-w-[150px]">{fileName}</span>}
              </div>
            </div>

            {/* Scanner Overlay */}
            <AnimatePresence>
              {isScanning && (
                <motion.div 
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  className="fixed inset-0 z-[100] bg-black flex flex-col items-center justify-center"
                >
                  <div className="relative w-full h-full max-w-xl mx-auto overflow-hidden bg-[#000]">
                    <video 
                      ref={videoRef} 
                      className="absolute inset-0 w-full h-full object-cover z-0"
                      autoPlay
                      muted
                      playsInline
                    />
                    
                    {/* Camera Loading State */}
                    {!isCameraReady && (
                      <div className="absolute inset-0 flex items-center justify-center bg-black z-50">
                        <div className="flex flex-col items-center gap-4">
                          <RefreshCcw className="w-10 h-10 text-blue-500 animate-spin" />
                          <p className="text-white text-sm font-medium">Conectando câmera...</p>
                        </div>
                      </div>
                    )}
                    
                    {/* Scanner UI Elements */}
                    <div className="absolute inset-0 flex flex-col items-center justify-center z-20">
                      <div className="w-64 h-64 border-2 border-white/50 rounded-3xl relative overflow-hidden">
                        <div className="absolute inset-0 border-2 border-blue-500 rounded-3xl animate-pulse" />
                        <div className="absolute top-1/2 left-0 w-full h-0.5 bg-blue-500 shadow-[0_0_15px_rgba(59,130,246,0.8)] animate-[scan_2s_ease-in-out_infinite]" />
                      </div>
                      <p className="mt-8 text-white font-medium bg-black/40 px-4 py-2 rounded-full backdrop-blur-sm">
                        Aponte para o código de barras
                      </p>
                      {/* Debug info */}
                      <p className="mt-2 text-[8px] text-white/30 font-mono">
                        Status: {isCameraReady ? 'Ativa' : 'Iniciando'} | Tracks: {videoRef.current?.srcObject ? (videoRef.current.srcObject as MediaStream).getVideoTracks().length : 0}
                      </p>
                    </div>

                    <button
                      onClick={stopScanner}
                      className="absolute top-6 right-6 bg-white/10 hover:bg-white/20 p-3 rounded-full backdrop-blur-md text-white transition-all z-30"
                    >
                      <X className="w-6 h-6" />
                    </button>

                    <button
                      onClick={() => { stopScanner(); setTimeout(() => setIsScanning(true), 100); }}
                      className="absolute top-6 left-6 bg-white/10 hover:bg-white/20 p-3 rounded-full backdrop-blur-md text-white transition-all z-30"
                      title="Reiniciar Câmera"
                    >
                      <RefreshCcw className="w-6 h-6" />
                    </button>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            {/* Results Area */}
            <div className="min-h-[300px]">
              <AnimatePresence mode="wait">
                {result ? (
                  <motion.div
                    key="result"
                    initial={{ opacity: 0, scale: 0.95 }}
                    animate={{ opacity: 1, scale: 1 }}
                    exit={{ opacity: 0, scale: 0.95 }}
                    className="bg-white rounded-2xl border border-gray-200 shadow-xl overflow-hidden"
                  >
                    <div className="bg-blue-600 p-6 text-white">
                      <div className="flex items-center gap-2 opacity-80 mb-1">
                        <Hash className="w-4 h-4" />
                        <span className="text-xs font-mono tracking-widest uppercase">Código: {result.Código}</span>
                      </div>
                      <h3 className="text-2xl font-bold leading-tight">{result.Modelo}</h3>
                    </div>
                    
                    <div className="p-6 grid grid-cols-2 gap-4">
                      <div className="space-y-1">
                        <div className="flex items-center gap-1.5 text-gray-400">
                          <MapPin className="w-4 h-4" />
                          <span className="text-[10px] uppercase font-bold tracking-wider">Localização</span>
                        </div>
                        <p className="text-lg font-semibold text-gray-900">{result.Local}</p>
                      </div>
                      
                      <div className="space-y-1">
                        <div className="flex items-center gap-1.5 text-gray-400">
                          <Package className="w-4 h-4" />
                          <span className="text-[10px] uppercase font-bold tracking-wider">Estoque</span>
                        </div>
                        <p className={cn(
                          "text-lg font-bold",
                          Number(result.Quantidade) > 0 ? "text-emerald-600" : "text-red-500"
                        )}>
                          {result.Quantidade} unidades
                        </p>
                      </div>
                    </div>
                    
                    <div className="px-6 py-4 bg-gray-50 border-t border-gray-100 flex justify-center">
                      <button 
                        onClick={clearSearch}
                        className="text-sm font-semibold text-blue-600 hover:text-blue-700"
                      >
                        Nova Busca
                      </button>
                    </div>
                  </motion.div>
                ) : error ? (
                  <motion.div
                    key="error"
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    className="bg-red-50 border border-red-100 rounded-2xl p-8 text-center space-y-3"
                  >
                    <AlertCircle className="w-10 h-10 text-red-500 mx-auto" />
                    <div className="space-y-1">
                      <h4 className="font-bold text-red-900">{error}</h4>
                      <p className="text-red-600/70 text-sm">Tente novamente ou verifique o código digitado.</p>
                    </div>
                    <button 
                      onClick={clearSearch}
                      className="bg-white text-red-600 px-4 py-2 rounded-lg text-sm font-bold border border-red-200 shadow-sm"
                    >
                      Limpar
                    </button>
                  </motion.div>
                ) : (
                  <motion.div
                    key="empty"
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    className="h-[300px] flex flex-col items-center justify-center text-gray-400 space-y-4 border-2 border-dashed border-gray-200 rounded-2xl"
                  >
                    <div className="bg-gray-50 p-4 rounded-full">
                      <Camera className="w-8 h-8 opacity-20" />
                    </div>
                    <p className="text-sm font-medium">Aguardando leitura ou busca...</p>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          </div>
        )}
      </main>

      {/* Footer Info */}
      <footer className="max-w-xl mx-auto px-4 py-8 text-center">
        <p className="text-[10px] text-gray-400 uppercase tracking-[0.2em] font-bold">
          Sistema de Inventário Local • v1.0
        </p>
      </footer>

      <style>{`
        @keyframes scan {
          0%, 100% { transform: translateY(-32px); opacity: 0; }
          50% { transform: translateY(32px); opacity: 1; }
        }
      `}</style>
    </div>
  );
}
