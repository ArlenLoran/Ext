import React, { useState, useEffect, useMemo } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import JSZip from 'jszip';
import { 
  FileText, 
  Download, 
  Trash2, 
  Eye, 
  Search, 
  RefreshCw, 
  AlertCircle, 
  CheckCircle2, 
  X, 
  Loader2,
  FileSearch,
  ExternalLink,
  ChevronRight,
  Clock,
  HardDrive,
  ChevronLeft,
  CheckSquare,
  Square,
  Files
} from 'lucide-react';
import { 
  listPdfFilesFromFolder, 
  downloadFileFromSharePoint, 
  deleteFileFromSharePoint 
} from './services/sharepointService';
import { PdfFile } from './types';

export default function App() {
  const [files, setFiles] = useState<PdfFile[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [viewingFile, setViewingFile] = useState<PdfFile | null>(null);
  const [pdfUrl, setPdfUrl] = useState<string | null>(null);
  const [notification, setNotification] = useState<{ type: 'success' | 'error'; message: string } | null>(null);
  const [isDeleting, setIsDeleting] = useState<string | null>(null);
  const [currentPage, setCurrentPage] = useState(1);
  const [selectedFiles, setSelectedFiles] = useState<Set<string>>(new Set());
  const [isDownloadingZip, setIsDownloadingZip] = useState(false);
  const isDevMode = !window._spPageContextInfo;

  const itemsPerPage = 15;
  const folderPath = 'Shared Documents/DACE';

  const fetchFiles = async () => {
    setLoading(true);
    setError(null);
    try {
      const pdfFiles = await listPdfFilesFromFolder(folderPath);
      setFiles(pdfFiles);
    } catch (err: any) {
      console.error('Erro ao buscar PDFs:', err);
      setError(err.message || 'Erro ao conectar com o SharePoint.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchFiles();
  }, []);

  useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm]);

  const showNotification = (type: 'success' | 'error', message: string) => {
    setNotification({ type, message });
    setTimeout(() => setNotification(null), 5000);
  };

  const handleView = async (file: PdfFile) => {
    try {
      const blob = await downloadFileFromSharePoint(file.serverRelativeUrl, file.name);
      const url = URL.createObjectURL(blob);
      setPdfUrl(url);
      setViewingFile(file);
    } catch (err) {
      showNotification('error', 'Não foi possível carregar o PDF para visualização.');
    }
  };

  const handleDownload = async (file: PdfFile) => {
    try {
      const blob = await downloadFileFromSharePoint(file.serverRelativeUrl, file.name);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = file.name;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      showNotification('success', `Download de ${file.name} iniciado.`);
    } catch (err) {
      showNotification('error', 'Falha ao baixar o arquivo.');
    }
  };

  const handleZipDownload = async () => {
    if (selectedFiles.size === 0) return;
    
    setIsDownloadingZip(true);
    try {
      const zip = new JSZip();
      const selectedPdfFiles = files.filter(f => selectedFiles.has(f.serverRelativeUrl));
      
      const downloadPromises = selectedPdfFiles.map(async (file) => {
        const blob = await downloadFileFromSharePoint(file.serverRelativeUrl, file.name);
        zip.file(file.name, blob);
      });

      await Promise.all(downloadPromises);
      
      const content = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(content);
      const a = document.createElement('a');
      a.href = url;
      a.download = `DACE_Batch_${new Date().getTime()}.zip`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      
      showNotification('success', `${selectedFiles.size} arquivos baixados em um .zip`);
      setSelectedFiles(new Set());
    } catch (err) {
      console.error('Erro ao gerar ZIP:', err);
      showNotification('error', 'Falha ao gerar o arquivo .zip.');
    } finally {
      setIsDownloadingZip(false);
    }
  };

  const toggleSelectAll = () => {
    if (selectedFiles.size === filteredFiles.length) {
      setSelectedFiles(new Set());
    } else {
      setSelectedFiles(new Set(filteredFiles.map(f => f.serverRelativeUrl)));
    }
  };

  const toggleSelectFile = (url: string) => {
    const next = new Set(selectedFiles);
    if (next.has(url)) {
      next.delete(url);
    } else {
      next.add(url);
    }
    setSelectedFiles(next);
  };

  const handleDelete = async (file: PdfFile) => {
    if (!window.confirm(`Tem certeza que deseja excluir o arquivo "${file.name}"?`)) return;

    setIsDeleting(file.serverRelativeUrl);
    try {
      await deleteFileFromSharePoint(file.serverRelativeUrl);
      setFiles(prev => prev.filter(f => f.serverRelativeUrl !== file.serverRelativeUrl));
      showNotification('success', 'Arquivo excluído com sucesso.');
    } catch (err) {
      showNotification('error', 'Erro ao excluir o arquivo.');
    } finally {
      setIsDeleting(null);
    }
  };

  const filteredFiles = useMemo(() => {
    return files.filter(f => f.name.toLowerCase().includes(searchTerm.toLowerCase()));
  }, [files, searchTerm]);

  const totalPages = Math.ceil(filteredFiles.length / itemsPerPage);
  const paginatedFiles = useMemo(() => {
    const start = (currentPage - 1) * itemsPerPage;
    return filteredFiles.slice(start, start + itemsPerPage);
  }, [filteredFiles, currentPage]);

  const formatSize = (bytes: number) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const formatDate = (dateStr: string) => {
    return new Date(dateStr).toLocaleString('pt-BR');
  };

  return (
    <div className="min-h-screen font-sans text-dhl-dark bg-gray-50">
      {/* Header */}
      <header className="bg-dhl-red text-white py-4 px-6 shadow-lg flex items-center justify-between sticky top-0 z-50">
        <div className="flex items-center gap-3">
          <div className="bg-dhl-yellow p-2 rounded-sm">
            <FileSearch className="text-dhl-red w-8 h-8" />
          </div>
          <div>
            <h1 className="text-2xl font-black tracking-tighter italic">DHL <span className="text-dhl-yellow not-italic font-bold ml-1">DACE MANAGER</span></h1>
            <div className="flex items-center gap-2">
              <p className="text-xs opacity-80 uppercase tracking-widest font-semibold">Gestão de Notas Fiscais PDF</p>
              {isDevMode && (
                <span className="bg-white/20 text-[8px] px-1.5 py-0.5 rounded border border-white/30 font-black uppercase tracking-tighter">Dev Mode (Mock Data)</span>
              )}
            </div>
          </div>
        </div>
        <div className="hidden md:block text-right">
          <p className="text-sm font-bold">SharePoint Integration</p>
          <p className="text-xs opacity-70">Pasta: {folderPath}</p>
        </div>
      </header>

      <main className="max-w-6xl mx-auto p-6 space-y-6">
        {/* Dev Mode Banner */}
        {isDevMode && (
          <motion.div 
            initial={{ opacity: 0, y: -10 }}
            animate={{ opacity: 1, y: 0 }}
            className="bg-blue-50 border border-blue-100 p-4 rounded-2xl flex items-start gap-3"
          >
            <AlertCircle className="text-blue-500 shrink-0 w-5 h-5" />
            <div>
              <p className="text-xs font-bold text-blue-800 uppercase tracking-tight">Ambiente de Desenvolvimento (AI Studio)</p>
              <p className="text-[11px] text-blue-600 mt-0.5">
                O contexto do SharePoint não foi detectado. O sistema está exibindo **dados fictícios (Mock Data)** para que você possa visualizar a interface e testar as funcionalidades. Quando publicado no SharePoint, o sistema conectará automaticamente à pasta real.
              </p>
            </div>
          </motion.div>
        )}

        {/* Notification Toast */}
        <AnimatePresence>
          {notification && (
            <motion.div
              initial={{ opacity: 0, y: -50 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9 }}
              className={`fixed top-24 right-6 z-[300] p-4 rounded-lg shadow-2xl flex items-center gap-3 border-l-4 ${
                notification.type === 'success' ? 'bg-white border-green-500 text-green-800' : 'bg-white border-red-500 text-red-800'
              }`}
            >
              {notification.type === 'success' ? <CheckCircle2 className="text-green-500" /> : <XCircle className="text-red-500" />}
              <p className="font-bold text-sm">{notification.message}</p>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Search and Actions Bar */}
        <div className="flex flex-col md:flex-row gap-4 items-center justify-between bg-white p-4 rounded-2xl shadow-sm border border-gray-100">
          <div className="relative w-full md:w-96">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 w-5 h-5" />
            <input
              type="text"
              placeholder="Buscar por nome do arquivo..."
              className="w-full pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-xl focus:ring-2 focus:ring-dhl-red focus:border-transparent outline-none transition-all"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          
          <div className="flex items-center gap-3 w-full md:w-auto">
            {selectedFiles.size > 0 && (
              <button
                onClick={handleZipDownload}
                disabled={isDownloadingZip}
                className="flex items-center justify-center gap-2 px-4 py-2 bg-dhl-yellow text-dhl-red hover:bg-yellow-400 rounded-xl transition-all font-bold text-sm w-full md:w-auto disabled:opacity-50 shadow-sm"
              >
                {isDownloadingZip ? (
                  <Loader2 className="w-4 h-4 animate-spin" />
                ) : (
                  <Files className="w-4 h-4" />
                )}
                ZIP ({selectedFiles.size})
              </button>
            )}
            <button
              onClick={fetchFiles}
              disabled={loading}
              className="flex items-center justify-center gap-2 px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-xl transition-all font-bold text-sm w-full md:w-auto disabled:opacity-50"
            >
              <RefreshCw className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} />
              Atualizar
            </button>
            <div className="bg-dhl-red/10 px-4 py-2 rounded-xl border border-dhl-red/20">
              <span className="text-dhl-red font-black text-sm uppercase tracking-widest">
                {filteredFiles.length} Arquivos
              </span>
            </div>
          </div>
        </div>

        {/* Main Content Area */}
        <div className="bg-white rounded-3xl shadow-sm border border-gray-100 overflow-hidden">
          {loading && files.length === 0 ? (
            <div className="flex flex-col items-center justify-center py-20 gap-4">
              <Loader2 className="w-12 h-12 text-dhl-red animate-spin" />
              <p className="text-gray-500 font-bold animate-pulse uppercase tracking-widest text-xs">Conectando ao SharePoint...</p>
            </div>
          ) : error ? (
            <div className="flex flex-col items-center justify-center py-20 gap-4 px-6 text-center">
              <div className="bg-red-50 p-4 rounded-full">
                <AlertCircle className="w-12 h-12 text-dhl-red" />
              </div>
              <h3 className="text-xl font-black text-dhl-dark italic uppercase">Erro de Conexão</h3>
              <p className="text-gray-500 max-w-md">{error}</p>
              <button 
                onClick={fetchFiles}
                className="mt-4 px-6 py-2 bg-dhl-red text-white font-bold rounded-xl hover:bg-red-700 transition-all shadow-lg shadow-red-200"
              >
                Tentar Novamente
              </button>
            </div>
          ) : filteredFiles.length === 0 ? (
            <div className="flex flex-col items-center justify-center py-20 gap-4">
              <div className="bg-gray-50 p-4 rounded-full">
                <FileText className="w-12 h-12 text-gray-300" />
              </div>
              <p className="text-gray-400 font-bold uppercase tracking-widest text-xs">Nenhum arquivo PDF encontrado</p>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-gray-50 border-b border-gray-100">
                    <th className="px-6 py-4 w-10">
                      <button 
                        onClick={toggleSelectAll}
                        className="text-gray-400 hover:text-dhl-red transition-colors"
                      >
                        {selectedFiles.size === filteredFiles.length && filteredFiles.length > 0 ? (
                          <CheckSquare className="w-5 h-5 text-dhl-red" />
                        ) : (
                          <Square className="w-5 h-5" />
                        )}
                      </button>
                    </th>
                    <th className="px-6 py-4 text-[10px] font-black uppercase tracking-widest text-gray-400">Arquivo</th>
                    <th className="px-6 py-4 text-[10px] font-black uppercase tracking-widest text-gray-400 hidden sm:table-cell">Data de Criação</th>
                    <th className="px-6 py-4 text-[10px] font-black uppercase tracking-widest text-gray-400 hidden md:table-cell">Tamanho</th>
                    <th className="px-6 py-4 text-[10px] font-black uppercase tracking-widest text-gray-400 text-right">Ações</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {paginatedFiles.map((file) => (
                    <motion.tr 
                      key={file.serverRelativeUrl}
                      initial={{ opacity: 0 }}
                      animate={{ opacity: 1 }}
                      className={`hover:bg-gray-50/50 transition-colors group ${selectedFiles.has(file.serverRelativeUrl) ? 'bg-red-50/30' : ''}`}
                    >
                      <td className="px-6 py-4">
                        <button 
                          onClick={() => toggleSelectFile(file.serverRelativeUrl)}
                          className="text-gray-400 hover:text-dhl-red transition-colors"
                        >
                          {selectedFiles.has(file.serverRelativeUrl) ? (
                            <CheckSquare className="w-5 h-5 text-dhl-red" />
                          ) : (
                            <Square className="w-5 h-5" />
                          )}
                        </button>
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex items-center gap-3">
                          <div className="bg-red-50 p-2 rounded-lg group-hover:bg-red-100 transition-colors">
                            <FileText className="w-5 h-5 text-dhl-red" />
                          </div>
                          <div>
                            <p className="text-sm font-bold text-dhl-dark truncate max-w-[200px] sm:max-w-xs" title={file.name}>
                              {file.name}
                            </p>
                            <p className="text-[10px] text-gray-400 font-medium uppercase tracking-tighter sm:hidden">
                              {formatDate(file.timeCreated)} • {formatSize(file.size)}
                            </p>
                          </div>
                        </div>
                      </td>
                      <td className="px-6 py-4 hidden sm:table-cell">
                        <div className="flex items-center gap-2 text-gray-500">
                          <Clock size={14} className="opacity-50" />
                          <span className="text-xs font-medium">{formatDate(file.timeCreated)}</span>
                        </div>
                      </td>
                      <td className="px-6 py-4 hidden md:table-cell">
                        <div className="flex items-center gap-2 text-gray-500">
                          <HardDrive size={14} className="opacity-50" />
                          <span className="text-xs font-medium">{formatSize(file.size)}</span>
                        </div>
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex items-center justify-end gap-2">
                          <button
                            onClick={() => handleView(file)}
                            className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-all"
                            title="Visualizar"
                          >
                            <Eye size={18} />
                          </button>
                          <button
                            onClick={() => handleDownload(file)}
                            className="p-2 text-green-600 hover:bg-green-50 rounded-lg transition-all"
                            title="Download"
                          >
                            <Download size={18} />
                          </button>
                          <button
                            onClick={() => handleDelete(file)}
                            disabled={isDeleting === file.serverRelativeUrl}
                            className="p-2 text-dhl-red hover:bg-red-50 rounded-lg transition-all disabled:opacity-50"
                            title="Excluir"
                          >
                            {isDeleting === file.serverRelativeUrl ? (
                              <Loader2 size={18} className="animate-spin" />
                            ) : (
                              <Trash2 size={18} />
                            )}
                          </button>
                        </div>
                      </td>
                    </motion.tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Pagination Controls */}
          {!loading && filteredFiles.length > 0 && (
            <div className="px-6 py-4 bg-gray-50 border-t border-gray-100 flex flex-col sm:flex-row items-center justify-between gap-4">
              <p className="text-xs font-bold text-gray-400 uppercase tracking-widest">
                Mostrando {Math.min(filteredFiles.length, (currentPage - 1) * itemsPerPage + 1)} - {Math.min(filteredFiles.length, currentPage * itemsPerPage)} de {filteredFiles.length}
              </p>
              
              <div className="flex items-center gap-1">
                <button
                  onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                  disabled={currentPage === 1}
                  className="p-2 text-gray-500 hover:bg-white hover:shadow-sm rounded-lg transition-all disabled:opacity-30"
                >
                  <ChevronLeft size={18} />
                </button>
                
                <div className="flex items-center gap-1">
                  {Array.from({ length: totalPages }, (_, i) => i + 1)
                    .filter(p => {
                      if (totalPages <= 5) return true;
                      if (p === 1 || p === totalPages) return true;
                      return Math.abs(p - currentPage) <= 1;
                    })
                    .map((p, i, arr) => (
                      <React.Fragment key={p}>
                        {i > 0 && arr[i-1] !== p - 1 && <span className="text-gray-300 px-1">...</span>}
                        <button
                          onClick={() => setCurrentPage(p)}
                          className={`w-8 h-8 flex items-center justify-center rounded-lg text-xs font-black transition-all ${
                            currentPage === p 
                              ? 'bg-dhl-red text-white shadow-md shadow-red-200' 
                              : 'text-gray-500 hover:bg-white hover:shadow-sm'
                          }`}
                        >
                          {p}
                        </button>
                      </React.Fragment>
                    ))
                  }
                </div>

                <button
                  onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                  disabled={currentPage === totalPages}
                  className="p-2 text-gray-500 hover:bg-white hover:shadow-sm rounded-lg transition-all disabled:opacity-30"
                >
                  <ChevronRight size={18} />
                </button>
              </div>
            </div>
          )}
        </div>
      </main>

      {/* PDF Viewer Modal */}
      <AnimatePresence>
        {viewingFile && pdfUrl && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 md:p-8">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => {
                setViewingFile(null);
                setPdfUrl(null);
              }}
              className="absolute inset-0 bg-dhl-dark/80 backdrop-blur-sm"
            />
            <motion.div
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-5xl h-full bg-white rounded-3xl shadow-2xl overflow-hidden flex flex-col"
            >
              <div className="bg-white border-b border-gray-100 p-4 flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="bg-red-50 p-2 rounded-lg">
                    <FileText className="text-dhl-red w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-sm font-black text-dhl-dark truncate max-w-[200px] sm:max-w-md">{viewingFile.name}</h3>
                    <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Visualização de Documento</p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => handleDownload(viewingFile)}
                    className="p-2 text-gray-500 hover:bg-gray-100 rounded-xl transition-all"
                    title="Baixar"
                  >
                    <Download size={20} />
                  </button>
                  <button
                    onClick={() => {
                      setViewingFile(null);
                      setPdfUrl(null);
                    }}
                    className="p-2 text-gray-500 hover:bg-gray-100 rounded-xl transition-all"
                  >
                    <X size={20} />
                  </button>
                </div>
              </div>
              <div className="flex-1 bg-gray-100 relative">
                <iframe
                  src={`${pdfUrl}#toolbar=0`}
                  className="w-full h-full border-none"
                  title="PDF Viewer"
                />
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Footer */}
      <footer className="max-w-6xl mx-auto p-6 text-center">
        <p className="text-[10px] font-black text-gray-300 uppercase tracking-[0.3em]">
          DHL DACE Manager • Excellence. Simply Delivered.
        </p>
      </footer>
    </div>
  );
}

function XCircle(props: any) {
  return (
    <svg
      {...props}
      xmlns="http://www.w3.org/2000/svg"
      width="24"
      height="24"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <circle cx="12" cy="12" r="10" />
      <path d="m15 9-6 6" />
      <path d="m9 9 6 6" />
    </svg>
  );
}
