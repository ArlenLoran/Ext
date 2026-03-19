/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useEffect, useMemo, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Upload, 
  FileText, 
  CheckCircle2, 
  XCircle, 
  AlertCircle, 
  ChevronDown, 
  ChevronUp,
  Trash2,
  FileSearch,
  Mail,
  Loader2,
  Settings,
  Plus,
  Edit2,
  Save,
  X,
  Braces,
  ShieldAlert,
  Search,
  ChevronLeft,
  ChevronRight,
  History,
  RotateCcw,
  Clock,
  Calendar,
  Download,
  ArrowRight,
  RefreshCw,
  Menu,
  MoreVertical
} from 'lucide-react';
import { sendEmail, buildXmlDivergenceEmailHtml, buildBatchXmlDivergenceEmailHtml } from './services/emailService';
import { listXmlFilesFromFolder, renameXmlFileAsValidated, revertXmlFileValidation, downloadFileFromSharePoint, listAllXmlFilesFromFolder } from './services/sharepointService';
import { SharePointListsService } from './services/sharepointLists';

import { useNotifications } from './hooks/useNotifications';
import { useXMLValidator } from './hooks/useXMLValidator';
import { useSharePointManager } from './hooks/useSharePointManager';
import { useResults } from './hooks/useResults';
import { ValidationResult } from './types';

export default function App() {
  const { notification, showNotification, setNotification } = useNotifications();
  const { results, setResults, expandedIndices, setExpandedIndices, resultsFilters, setResultsFilters, clearAll, toggleExpand, removeResult } = useResults();
  const { mandatoryTags, setMandatoryTags, osForbiddenPatterns, setOsForbiddenPatterns, registeredProducts, setRegisteredProducts, validateXML, extractXmlMetadata, checkNtvStatus, checkOsStatus } = useXMLValidator();
  
  const [recipients, setRecipients] = useState<string[]>(() => {
    const saved = localStorage.getItem('dhl_recipients');
    return saved ? JSON.parse(saved) : [];
  });
  const [newEmail, setNewEmail] = useState('');
  const [editingEmail, setEditingEmail] = useState<{ index: number; value: string } | null>(null);
  const [emailSearch, setEmailSearch] = useState('');
  const [emailPage, setEmailPage] = useState(1);

  const [newTagName, setNewTagName] = useState('');
  const [newTagRef, setNewTagRef] = useState('');
  const [newPattern, setNewPattern] = useState('');
  const [newProduct, setNewProduct] = useState('');

  const {
    isSpAvailable, setIsSpAvailable,
    isSpInitialized, setIsSpInitialized,
    isInitializingSp, setIsInitializingSp,
    isFetchingSharePoint, setIsFetchingSharePoint,
    revalidationItems, setRevalidationItems,
    showRevalidation, setShowRevalidation,
    isFetchingRevalidation,
    revalidationSearch, setRevalidationSearch,
    revalidationPage, setRevalidationPage,
    revalidationStartDate, setRevalidationStartDate,
    revalidationEndDate, setRevalidationEndDate,
    fullHistory, setFullHistory,
    showFullHistory, setShowFullHistory,
    isFetchingFullHistory,
    fullHistorySearch, setFullHistorySearch,
    fullHistoryPage, setFullHistoryPage,
    fullHistoryStartDate, setFullHistoryStartDate,
    fullHistoryEndDate, setFullHistoryEndDate,
    spStats,
    spFilesList, setSpFilesList,
    isFetchingSpStats,
    showSpManager, setShowSpManager,
    spManagerSearch, setSpManagerSearch,
    spManagerPage, setSpManagerPage,
    spManagerStartDate, setSpManagerStartDate,
    spManagerEndDate, setSpManagerEndDate,
    filteredSpFiles,
    fetchSpStats,
    checkSpInitialization,
    loadRevalidationFromSharePoint,
    loadFullHistoryFromSharePoint,
    downloadFromSharePoint,
    handleRevertSpFile,
    handleRevertValidation,
    initializeSharePoint,
    syncAllToSharePoint
  } = useSharePointManager(
    showNotification,
    recipients,
    setRecipients,
    mandatoryTags,
    setMandatoryTags,
    osForbiddenPatterns,
    setOsForbiddenPatterns
  );

  const [isDragging, setIsDragging] = useState(false);
  const [sendingEmailIdx, setSendingEmailIdx] = useState<number | null>(null);
  const [isSendingBatch, setIsSendingBatch] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showMenu, setShowMenu] = useState(false);
  const menuRef = useRef<HTMLDivElement>(null);
  const [tagSearch, setTagSearch] = useState('');
  const [tagPage, setTagPage] = useState(1);
  const [patternSearch, setPatternSearch] = useState('');
  const [patternPage, setPatternPage] = useState(1);
  const [productSearch, setProductSearch] = useState('');
  const [productPage, setProductPage] = useState(1);
  const itemsPerPage = 10;

  // Check SharePoint Context on Mount
  React.useEffect(() => {
    const available = SharePointListsService.isContextAvailable();
    setIsSpAvailable(available);
    if (available) {
      checkSpInitialization();
      fetchSpStats();
      // Refresh stats every 5 minutes
      const interval = setInterval(fetchSpStats, 5 * 60 * 1000);
      return () => clearInterval(interval);
    }
  }, [checkSpInitialization, fetchSpStats, setIsSpAvailable]);

  const handleSharePointImport = async (filesOrEvent?: any[] | React.MouseEvent) => {
    let filesToProcess: any[] = [];
    
    if (Array.isArray(filesOrEvent)) {
      filesToProcess = filesOrEvent;
    } else {
      // Se for um evento ou chamado sem argumentos, busca os arquivos pendentes no SharePoint
      setIsFetchingSharePoint(true);
      try {
        const fetchedFiles = await listXmlFilesFromFolder('SiteAssets/XMLs');
        if (fetchedFiles.length === 0) {
          showNotification('success', 'Nenhum arquivo XML pendente no SharePoint.');
          setIsFetchingSharePoint(false);
          return;
        }
        filesToProcess = fetchedFiles;
      } catch (error) {
        console.error('Erro ao buscar arquivos do SharePoint:', error);
        showNotification('error', 'Erro ao buscar arquivos do SharePoint.');
        setIsFetchingSharePoint(false);
        return;
      }
    }

    if (filesToProcess.length === 0) return;
    
    setIsFetchingSharePoint(true);
    try {
      const spUrlMap: Record<string, string> = {};
      const importedFiles: File[] = [];

      for (const file of filesToProcess) {
        try {
          // Se já tivermos o objeto File (vindo de listXmlFilesFromFolder), usamos ele
          // Caso contrário, baixamos o arquivo
          let xmlFile: File;
          if (file.file instanceof File) {
            xmlFile = file.file;
          } else {
            const blob = await SharePointListsService.downloadFile(file.serverRelativeUrl);
            xmlFile = new File([blob], file.name, { type: 'text/xml' });
          }
          
          const fileName = file.name;
          const serverRelativeUrl = file.serverRelativeUrl;

          // Rename the file if it's valid
          const res = await validateXML(xmlFile);
          if (res.isValid) {
            const newName = fileName.replace(/\.xml$/i, ' validado.xml');
            const renamedFile = new File([await xmlFile.arrayBuffer()], newName, { type: 'text/xml' });
            
            // Move/Rename in SharePoint
            const folderPath = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));
            const newUrl = `${folderPath}/${newName}`;
            await SharePointListsService.moveFile(serverRelativeUrl, newUrl);
            
            spUrlMap[newName] = newUrl;
            importedFiles.push(renamedFile);
          } else {
            spUrlMap[fileName] = serverRelativeUrl;
            importedFiles.push(xmlFile);
          }
        } catch (err) {
          console.error(`Erro ao importar arquivo ${file.name}:`, err);
        }
      }

      if (importedFiles.length > 0) {
        await handleFiles(importedFiles, spUrlMap);
        showNotification('success', `${importedFiles.length} arquivo(s) importado(s) e validado(s) do SharePoint.`);
      }
    } catch (error) {
      console.error('Erro ao importar do SharePoint:', error);
      showNotification('error', 'Falha ao importar arquivos do SharePoint.');
    } finally {
      setIsFetchingSharePoint(false);
      fetchSpStats(); // Refresh stats after import/rename
    }
  };


  React.useEffect(() => {
    localStorage.setItem('dhl_recipients', JSON.stringify(recipients));
  }, [recipients]);

  React.useEffect(() => {
    localStorage.setItem('dhl_mandatory_tags', JSON.stringify(mandatoryTags));
  }, [mandatoryTags]);

  React.useEffect(() => {
    localStorage.setItem('dhl_os_forbidden_patterns', JSON.stringify(osForbiddenPatterns));
  }, [osForbiddenPatterns]);

  React.useEffect(() => {
    localStorage.setItem('dhl_registered_products', JSON.stringify(registeredProducts));
  }, [registeredProducts]);

  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(event.target as Node)) {
        setShowMenu(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const addRecipient = (e: React.FormEvent) => {
    e.preventDefault();
    const email = newEmail.trim();
    if (!email || !email.includes('@')) {
      setNotification({ type: 'error', message: 'E-mail inválido.' });
      return;
    }
    if (recipients.includes(email)) {
      setNotification({ type: 'error', message: 'E-mail já cadastrado.' });
      return;
    }
    setRecipients([...recipients, email]);
    if (isSpInitialized) {
      SharePointListsService.createItem('DHL_Recipients', { Title: email });
    }
    setNewEmail('');
    setNotification({ type: 'success', message: 'E-mail adicionado com sucesso!' });
    setTimeout(() => setNotification(null), 3000);
  };

  const removeRecipient = async (index: number) => {
    const emailToRemove = recipients[index];
    setRecipients(recipients.filter((_, i) => i !== index));
    if (isSpInitialized) {
      try {
        const items = await SharePointListsService.getItemsByFilter('DHL_Recipients', `Title eq '${emailToRemove}'`, { select: ['Id'] });
        if (items.length > 0) {
          await SharePointListsService.deleteItem('DHL_Recipients', items[0].Id);
        }
      } catch (error) {
        console.error('Erro ao remover do SharePoint:', error);
      }
    }
  };

  const startEdit = (index: number, value: string) => {
    setEditingEmail({ index, value });
  };

  const saveEdit = async () => {
    if (!editingEmail) return;
    const oldEmail = recipients[editingEmail.index];
    const email = editingEmail.value.trim();
    if (!email || !email.includes('@')) {
      setNotification({ type: 'error', message: 'E-mail inválido.' });
      return;
    }
    const updated = [...recipients];
    updated[editingEmail.index] = email;
    setRecipients(updated);

    if (isSpInitialized) {
      try {
        const items = await SharePointListsService.getItemsByFilter('DHL_Recipients', `Title eq '${oldEmail}'`, { select: ['Id'] });
        if (items.length > 0) {
          await SharePointListsService.updateItem('DHL_Recipients', items[0].Id, { Title: email });
        }
      } catch (error) {
        console.error('Erro ao atualizar no SharePoint:', error);
      }
    }

    setEditingEmail(null);
  };

  const addTag = (e: React.FormEvent) => {
    e.preventDefault();
    const name = newTagName.trim();
    const tag = newTagRef.trim();
    if (!name || !tag) {
      setNotification({ type: 'error', message: 'Preencha Nome e Referência.' });
      return;
    }
    if (mandatoryTags.some(t => t.tag.toLowerCase() === tag.toLowerCase())) {
      setNotification({ type: 'error', message: 'Esta referência já existe.' });
      return;
    }
    setMandatoryTags([...mandatoryTags, { name, tag }]);
    if (isSpInitialized) {
      SharePointListsService.createItem('DHL_MandatoryTags', { Title: name, TagRef: tag });
    }
    setNewTagName('');
    setNewTagRef('');
    setNotification({ type: 'success', message: 'Campo obrigatório adicionado!' });
    setTimeout(() => setNotification(null), 3000);
  };

  const removeTag = async (tagRef: string) => {
    setMandatoryTags(mandatoryTags.filter(t => t.tag !== tagRef));
    if (isSpInitialized) {
      try {
        const items = await SharePointListsService.getItemsByFilter('DHL_MandatoryTags', `TagRef eq '${tagRef}'`, { select: ['Id'] });
        if (items.length > 0) {
          await SharePointListsService.deleteItem('DHL_MandatoryTags', items[0].Id);
        }
      } catch (error) {
        console.error('Erro ao remover tag do SharePoint:', error);
      }
    }
  };

  const addPattern = (e: React.FormEvent) => {
    e.preventDefault();
    const pattern = newPattern.trim();
    if (!pattern) return;
    try {
      new RegExp(pattern);
    } catch (e) {
      setNotification({ type: 'error', message: 'Regex inválido.' });
      return;
    }
    if (osForbiddenPatterns.includes(pattern)) {
      setNotification({ type: 'error', message: 'Padrão já existe.' });
      return;
    }
    setOsForbiddenPatterns([...osForbiddenPatterns, pattern]);
    if (isSpInitialized) {
      SharePointListsService.createItem('DHL_OSPatterns', { Title: pattern });
    }
    setNewPattern('');
  };

  const removePattern = async (pattern: string) => {
    setOsForbiddenPatterns(osForbiddenPatterns.filter(p => p !== pattern));
    if (isSpInitialized) {
      try {
        const items = await SharePointListsService.getItemsByFilter('DHL_OSPatterns', `Title eq '${pattern}'`, { select: ['Id'] });
        if (items.length > 0) {
          await SharePointListsService.deleteItem('DHL_OSPatterns', items[0].Id);
        }
      } catch (error) {
        console.error('Erro ao remover padrão do SharePoint:', error);
      }
    }
  };

  const addProduct = (e: React.FormEvent) => {
    e.preventDefault();
    const product = newProduct.trim();
    if (!product) return;
    if (registeredProducts.includes(product)) {
      setNotification({ type: 'error', message: 'Produto já cadastrado.' });
      return;
    }
    setRegisteredProducts([...registeredProducts, product]);
    if (isSpInitialized) {
      SharePointListsService.createItem('DHL_RegisteredProducts', { Title: product });
    }
    setNewProduct('');
  };

  const removeProduct = async (product: string) => {
    setRegisteredProducts(registeredProducts.filter(p => p !== product));
    if (isSpInitialized) {
      try {
        const items = await SharePointListsService.getItemsByFilter('DHL_RegisteredProducts', `Title eq '${product}'`, { select: ['Id'] });
        if (items.length > 0) {
          await SharePointListsService.deleteItem('DHL_RegisteredProducts', items[0].Id);
        }
      } catch (error) {
        console.error('Erro ao remover produto do SharePoint:', error);
      }
    }
  };


  const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result as string;
        const base64 = result.split(",")[1];
        resolve(base64);
      };
      reader.onerror = () => reject(new Error("Erro ao converter arquivo para base64"));
      reader.readAsDataURL(file);
    });
  };

  const handleSendReport = async (result: ValidationResult, index: number) => {
    if (recipients.length === 0) {
      setNotification({ type: 'error', message: 'Nenhum destinatário cadastrado.' });
      return;
    }
    setSendingEmailIdx(index);
    try {
      const html = buildXmlDivergenceEmailHtml({
        fileName: result.fileName,
        nNF: result.nNF,
        cnpj: result.cnpj,
        errors: result.errors,
        appUrl: window.location.href
      });

      const attachments = [{
        Name: result.fileName,
        ContentBytes: await fileToBase64(result.originalFile)
      }];

      // Pass the array directly to let the service handle the separator (semicolon)
      await sendEmail(recipients, `Divergência XML: ${result.fileName}`, html, attachments);
      
      setNotification({ type: 'success', message: `Relatório enviado para ${recipients.length} destinatário(s)!` });
      
      // Mark as sent
      setResults(prev => {
        const updated = [...prev];
        updated[index] = { ...updated[index], sent: true };
        return updated;
      });
    } catch (error) {
      console.error(error);
      setNotification({ type: 'error', message: 'Falha ao enviar relatório com anexo.' });
    } finally {
      setSendingEmailIdx(null);
      setTimeout(() => setNotification(null), 5000);
    }
  };

  const handleSendBatchReport = async () => {
    if (recipients.length === 0) {
      setNotification({ type: 'error', message: 'Nenhum destinatário cadastrado.' });
      return;
    }
    const resultsWithErrors = results.filter(r => r.errors.length > 0);
    if (resultsWithErrors.length === 0) return;

    setIsSendingBatch(true);
    try {
      const html = buildBatchXmlDivergenceEmailHtml({
        results: resultsWithErrors.map(r => ({
          fileName: r.fileName,
          nNF: r.nNF,
          cnpj: r.cnpj,
          errors: r.errors
        })),
        appUrl: window.location.href
      });

      const attachments = await Promise.all(resultsWithErrors.map(async (r) => ({
        Name: r.fileName,
        ContentBytes: await fileToBase64(r.originalFile)
      })));

      // Pass the array directly to let the service handle the separator (semicolon)
      await sendEmail(
        recipients, 
        `Relatório de Divergências em Lote (${resultsWithErrors.length} arquivos)`, 
        html,
        attachments
      );
      
      setNotification({ type: 'success', message: `Relatório de lote enviado para ${recipients.length} destinatário(s)!` });
      
      // Mark all results with errors as sent
      setResults(prev => prev.map(r => r.errors.length > 0 ? { ...r, sent: true } : r));
    } catch (error) {
      console.error(error);
      setNotification({ type: 'error', message: 'Falha ao enviar relatório em lote com anexos.' });
    } finally {
      setIsSendingBatch(false);
      setTimeout(() => setNotification(null), 5000);
    }
  };

  const handleFiles = async (files: FileList | File[], spUrlMap?: Record<string, string>): Promise<ValidationResult[]> => {
    const newResults: ValidationResult[] = [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i] instanceof File ? (files[i] as File) : (files[i] as any);
      if (file.type === "text/xml" || file.name.endsWith(".xml")) {
        const res = await validateXML(file);
        if (spUrlMap && spUrlMap[file.name]) {
          res.sharepointUrl = spUrlMap[file.name];
        }
        newResults.push(res);
      }
    }
    
    if (newResults.length > 0) {
      setResults(prev => [...newResults, ...prev]);

      // Trigger background NTV and OS checks for new results using fileName as unique key
      newResults.forEach((res) => {
        if (res.xProd) {
          checkNtvStatus(res.fileName, res.xProd, setResults);
        }
        if (res.osField && res.osField !== "Não encontrado") {
          checkOsStatus(res.fileName, res.osField, setResults);
        }
      });

      // Log to Full History if SharePoint context is available
      if (SharePointListsService.isContextAvailable()) {
        const userInfo = SharePointListsService.getUserInfo();
        
        // Use Promise.all to ensure all items are created before refreshing the list
        Promise.all(newResults.map(async (res) => {
          try {
            await SharePointListsService.createItem('DHL_FullHistory', {
              Title: res.fileName,
              Status: res.isValid ? 'Válido' : 'Inválido',
              nNF: res.nNF || '',
              CNPJ: res.cnpj || '',
              OS: res.osField || '',
              NCM: res.ncm || '',
              xProd: res.xProd || '',
              UserEmail: userInfo.email || 'Usuário Local',
              Source: res.sharepointUrl ? 'SharePoint' : 'Local',
              ValidationDate: new Date().toISOString()
            });
          } catch (err) {
            console.error('Erro ao logar no histórico completo:', err);
          }
        })).then(() => {
          // No automatic refresh as per user request
        });
      }
    }

    return newResults;
  };

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files) {
      handleFiles(e.dataTransfer.files);
    }
  }, []);

  const onDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const onDragLeave = () => {
    setIsDragging(false);
  };

  const downloadXml = (result: ValidationResult) => {
    const blob = new Blob([result.rawContent], { type: 'text/xml' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = result.fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const validateSpFileManually = async (spFile: { name: string; serverRelativeUrl: string }) => {
    try {
      showNotification('success', `Importando ${spFile.name}...`);
      const blob = await SharePointListsService.downloadFile(spFile.serverRelativeUrl);
      const file = new File([blob], spFile.name, { type: 'text/xml' });
      
      const spUrlMap = { [spFile.name]: spFile.serverRelativeUrl };
      await handleFiles([file], spUrlMap);
      
      setShowSpManager(false);
      showNotification('success', 'Arquivo importado para validação!');
      fetchSpStats(); // Refresh stats
    } catch (error) {
      console.error(error);
      showNotification('error', 'Erro ao importar arquivo do SharePoint.');
    }
  };

  return (
    <div className="min-h-screen font-sans text-dhl-dark">
      {/* Header */}
      <header className="bg-dhl-red text-white py-4 px-6 shadow-lg flex items-center justify-between sticky top-0 z-50">
        <div className="flex items-center gap-3">
          <div className="bg-dhl-yellow p-2 rounded-sm">
            <FileSearch className="text-dhl-red w-8 h-8" />
          </div>
          <div>
            <h1 className="text-2xl font-black tracking-tighter italic">DHL <span className="text-dhl-yellow not-italic font-bold ml-1">XML VALIDATOR</span></h1>
            <p className="text-xs opacity-80 uppercase tracking-widest font-semibold">Excellence. Simply Delivered.</p>
          </div>
        </div>
        <div className="hidden md:block text-right">
          <p className="text-sm font-bold">Logística de Documentos Fiscais</p>
          <p className="text-xs opacity-70">v1.0.4 - Produção</p>
        </div>
      </header>

      <main className="max-w-6xl mx-auto p-6 space-y-8">
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
              <button onClick={() => setNotification(null)} className="ml-4 opacity-50 hover:opacity-100">
                <Trash2 size={14} />
              </button>
            </motion.div>
          )}
        </AnimatePresence>

        {/* SharePoint Live Monitor Bar */}
        {isSpAvailable && (
          <motion.div 
            initial={{ opacity: 0, y: -10 }}
            animate={{ opacity: 1, y: 0 }}
            className="flex items-center justify-between bg-white/50 backdrop-blur-sm border border-gray-100 px-6 py-3 rounded-2xl shadow-sm"
          >
            <div className="flex items-center gap-8">
              <div className="flex items-center gap-3">
                <div className="relative">
                  <div className="w-2.5 h-2.5 rounded-full bg-blue-500" />
                  <div className="absolute inset-0 w-2.5 h-2.5 rounded-full bg-blue-500 animate-ping opacity-75" />
                </div>
                <span className="text-[10px] font-black uppercase tracking-widest text-gray-400">Monitor SharePoint</span>
              </div>
              
              <div className="flex items-center gap-6">
                <div className="flex items-center gap-2">
                  <span className="text-sm font-black text-blue-600">{spStats.analyzed}</span>
                  <span className="text-[9px] font-bold text-gray-400 uppercase tracking-tight">Analisados</span>
                </div>
                <div className="w-px h-4 bg-gray-200" />
                <div className="flex items-center gap-2">
                  <span className="text-sm font-black text-orange-500">{spStats.pending}</span>
                  <span className="text-[9px] font-bold text-gray-400 uppercase tracking-tight">Pendentes</span>
                </div>
              </div>
            </div>

            <button 
              onClick={() => setShowSpManager(true)}
              className="flex items-center gap-2 px-4 py-2 bg-white hover:bg-gray-50 border border-gray-200 rounded-xl transition-all group shadow-sm"
            >
              <span className="text-[10px] font-black uppercase tracking-widest text-gray-500 group-hover:text-dhl-dark">Gerenciar Pasta</span>
              <ChevronRight size={14} className="text-gray-400 group-hover:text-dhl-red transition-colors" />
            </button>
          </motion.div>
        )}

        {/* Hero / Upload Section */}
        <section className="space-y-4">
          <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 bg-white p-6 rounded-2xl shadow-sm border border-gray-100">
            <div className="flex items-center gap-4">
              <div className="bg-dhl-yellow p-3 rounded-xl shadow-inner">
                <FileText size={24} className="text-dhl-dark" />
              </div>
              <div>
                <h2 className="text-2xl font-black tracking-tighter text-dhl-dark italic uppercase leading-none">
                  Validação de Notas
                </h2>
                <div className="flex items-center gap-3 mt-2">
                  <div className="flex items-center gap-1.5 bg-green-50 px-2 py-0.5 rounded-full border border-green-100">
                    <span className="w-1.5 h-1.5 rounded-full bg-green-500" />
                    <span className="text-[9px] font-black text-green-700 uppercase tracking-widest">
                      {results.filter(r => r.errors.length === 0).length} OK
                    </span>
                  </div>
                  <div className="flex items-center gap-1.5 bg-red-50 px-2 py-0.5 rounded-full border border-red-100">
                    <span className="w-1.5 h-1.5 rounded-full bg-dhl-red" />
                    <span className="text-[9px] font-black text-dhl-red uppercase tracking-widest">
                      {results.filter(r => r.errors.length > 0).length} ERROS
                    </span>
                  </div>
                </div>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2 md:gap-3">
              {/* Actions Menu */}
              <div className="relative" ref={menuRef}>
                <button 
                  onClick={() => setShowMenu(!showMenu)}
                  className={`flex items-center gap-2 px-3 py-2 rounded-xl transition-all font-bold text-xs uppercase tracking-widest border shadow-sm ${showMenu ? 'bg-dhl-dark text-white border-dhl-dark' : 'bg-gray-50 text-gray-500 border-gray-200 hover:bg-gray-100'}`}
                >
                  <Menu size={16} />
                  <span className="hidden sm:inline">Opções</span>
                  <ChevronDown size={14} className={`transition-transform duration-200 ${showMenu ? 'rotate-180' : ''}`} />
                </button>

                <AnimatePresence>
                  {showMenu && (
                    <motion.div
                      initial={{ opacity: 0, y: 10, scale: 0.95 }}
                      animate={{ opacity: 1, y: 0, scale: 1 }}
                      exit={{ opacity: 0, y: 10, scale: 0.95 }}
                      className="absolute right-0 mt-2 w-56 bg-white rounded-2xl shadow-2xl border border-gray-100 py-2 z-50 overflow-hidden"
                    >
                      <div className="px-3 py-2 mb-1 border-b border-gray-50">
                        <span className="text-[9px] font-black uppercase tracking-[0.2em] text-gray-400">Menu de Ações</span>
                      </div>
                      
                      <button 
                        onClick={() => { setShowFullHistory(true); setShowMenu(false); }}
                        className="w-full flex items-center gap-3 px-4 py-2.5 text-left text-xs font-bold text-gray-600 hover:bg-gray-50 hover:text-dhl-dark transition-colors group"
                      >
                        <div className="p-1.5 bg-blue-50 rounded-lg group-hover:bg-blue-100 transition-colors">
                          <History size={14} className="text-blue-600" />
                        </div>
                        Histórico Completo
                      </button>

                      <button 
                        onClick={() => { setShowRevalidation(true); setShowMenu(false); }}
                        className="w-full flex items-center gap-3 px-4 py-2.5 text-left text-xs font-bold text-gray-600 hover:bg-gray-50 hover:text-dhl-dark transition-colors group"
                      >
                        <div className="p-1.5 bg-orange-50 rounded-lg group-hover:bg-orange-100 transition-colors">
                          <RotateCcw size={14} className="text-orange-600" />
                        </div>
                        Revalidação
                      </button>

                      <button 
                        onClick={() => { setShowSettings(!showSettings); setShowMenu(false); }}
                        className={`w-full flex items-center gap-3 px-4 py-2.5 text-left text-xs font-bold transition-colors group ${showSettings ? 'bg-dhl-dark/5 text-dhl-dark' : 'text-gray-600 hover:bg-gray-50 hover:text-dhl-dark'}`}
                      >
                        <div className={`p-1.5 rounded-lg transition-colors ${showSettings ? 'bg-dhl-dark text-white' : 'bg-gray-100 group-hover:bg-gray-200'}`}>
                          <Settings size={14} />
                        </div>
                        Configurações
                      </button>

                      {results.length > 0 && (
                        <>
                          <div className="my-1 border-t border-gray-50" />
                          <button 
                            onClick={() => { clearAll(); setShowMenu(false); }}
                            className="w-full flex items-center gap-3 px-4 py-2.5 text-left text-xs font-bold text-red-500 hover:bg-red-50 transition-colors group"
                          >
                            <div className="p-1.5 bg-red-50 rounded-lg group-hover:bg-red-100 transition-colors">
                              <Trash2 size={14} className="text-red-600" />
                            </div>
                            Limpar Tudo
                          </button>
                        </>
                      )}
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>

              {/* Primary Actions */}
              <button 
                onClick={handleSharePointImport}
                disabled={isFetchingSharePoint}
                className="bg-dhl-dark text-white px-4 py-2.5 rounded-xl transition-all flex items-center gap-2 font-black text-xs uppercase tracking-widest shadow-lg hover:bg-black disabled:opacity-50"
              >
                {isFetchingSharePoint ? (
                  <><Loader2 size={16} className="animate-spin" /> BUSCANDO...</>
                ) : (
                  <><FileSearch size={16} /> IMPORTAR SHAREPOINT</>
                )}
              </button>

              {results.some(r => r.errors.length > 0) && (
                <button 
                  onClick={handleSendBatchReport}
                  disabled={isSendingBatch || results.filter(r => r.errors.length > 0).every(r => r.sent)}
                  className="bg-dhl-red text-white px-4 py-2.5 rounded-xl transition-all flex items-center gap-2 font-black text-xs uppercase tracking-widest shadow-lg hover:bg-red-700 disabled:opacity-50"
                >
                  {isSendingBatch ? (
                    <><Loader2 size={16} className="animate-spin" /> ENVIANDO...</>
                  ) : results.filter(r => r.errors.length > 0).every(r => r.sent) ? (
                    <><CheckCircle2 size={16} /> LOTE ENVIADO</>
                  ) : (
                    <><Mail size={16} /> REPORTAR TUDO</>
                  )}
                </button>
              )}
            </div>
          </div>

          <AnimatePresence>
            {showSettings && (
              <motion.div 
                initial={{ height: 0, opacity: 0 }}
                animate={{ height: 'auto', opacity: 1 }}
                exit={{ height: 0, opacity: 0 }}
                className="overflow-hidden"
              >
                <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-8">
                  <div className="flex items-center justify-between">
                    <h3 className="text-xl font-black text-dhl-dark flex items-center gap-2 italic uppercase tracking-tighter">
                      <Settings size={24} className="text-dhl-red" /> Configurações do Sistema
                    </h3>
                    <button onClick={() => setShowSettings(false)} className="text-gray-400 hover:text-gray-600">
                      <X size={24} />
                    </button>
                  </div>

                  {/* SharePoint Integration Banner */}
                  <div className={`p-4 rounded-xl border flex items-center justify-between gap-4 ${isSpInitialized ? 'bg-green-50 border-green-100' : isSpAvailable ? 'bg-blue-50 border-blue-100' : 'bg-gray-50 border-gray-100'}`}>
                    <div className="flex items-center gap-3">
                      <div className={`p-2 rounded-lg ${isSpInitialized ? 'bg-green-100 text-green-600' : isSpAvailable ? 'bg-blue-100 text-blue-600' : 'bg-gray-200 text-gray-500'}`}>
                        {isSpInitialized ? <CheckCircle2 size={20} /> : <ShieldAlert size={20} />}
                      </div>
                      <div>
                        <h4 className={`text-sm font-black uppercase tracking-widest ${isSpInitialized ? 'text-green-700' : isSpAvailable ? 'text-blue-700' : 'text-gray-700'}`}>
                          Integração SharePoint
                        </h4>
                        <p className="text-xs text-gray-500">
                          {isSpInitialized 
                            ? 'As configurações estão sendo sincronizadas com as listas do SharePoint.' 
                            : isSpAvailable 
                              ? 'O contexto do SharePoint foi detectado. Clique ao lado para inicializar as listas de persistência.'
                              : 'Contexto do SharePoint não detectado. As configurações serão salvas apenas localmente.'}
                        </p>
                      </div>
                    </div>
                    {!isSpInitialized && (
                      <button 
                        onClick={isSpAvailable ? initializeSharePoint : () => {
                          const available = SharePointListsService.isContextAvailable();
                          setIsSpAvailable(available);
                          if (available) {
                            checkSpInitialization();
                            setNotification({ type: 'success', message: 'Contexto detectado!' });
                          } else {
                            setNotification({ type: 'error', message: 'Contexto ainda não encontrado.' });
                          }
                          setTimeout(() => setNotification(null), 3000);
                        }}
                        disabled={isInitializingSp}
                        className="bg-dhl-dark text-white px-4 py-2 rounded-lg text-xs font-black uppercase tracking-widest hover:bg-black transition-all flex items-center gap-2 disabled:opacity-50"
                      >
                        {isInitializingSp ? <Loader2 size={14} className="animate-spin" /> : <Plus size={14} />}
                        {isInitializingSp ? 'Inicializando...' : isSpAvailable ? 'Validar Contexto e Criar Listas' : 'Tentar Validar Contexto'}
                      </button>
                    )}
                    {isSpInitialized && (
                      <div className="flex items-center gap-2 text-green-600 font-black text-[10px] uppercase tracking-widest">
                        <CheckCircle2 size={14} /> Ativo
                      </div>
                    )}
                  </div>
                  
                  <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
                    {/* Recipients Section */}
                    <div className="space-y-4 flex flex-col h-full">
                      <div className="flex items-center justify-between">
                        <h4 className="text-sm font-black uppercase tracking-widest text-gray-400 flex items-center gap-2">
                          <Mail size={16} /> Destinatários
                        </h4>
                        <div className="relative">
                          <Search size={14} className="absolute left-2 top-1/2 -translate-y-1/2 text-gray-400" />
                          <input 
                            type="text"
                            placeholder="Filtrar..."
                            value={emailSearch}
                            onChange={(e) => { setEmailSearch(e.target.value); setEmailPage(1); }}
                            className="pl-7 pr-2 py-1 border border-gray-200 rounded text-[10px] focus:outline-none focus:ring-1 focus:ring-dhl-red/20 w-32"
                          />
                        </div>
                      </div>
                      
                      <form onSubmit={addRecipient} className="flex gap-2">
                        <input 
                          type="email" 
                          value={newEmail}
                          onChange={(e) => setNewEmail(e.target.value)}
                          placeholder="Novo e-mail..."
                          className="flex-1 px-3 py-2 border border-gray-300 rounded-md text-sm focus:outline-none focus:ring-2 focus:ring-dhl-red/20"
                        />
                        <button 
                          type="submit"
                          className="bg-dhl-dark text-white p-2 rounded-md hover:bg-black transition-colors"
                        >
                          <Plus size={16} />
                        </button>
                      </form>

                      <div className="flex-1 border border-gray-100 rounded-xl overflow-hidden bg-gray-50/50">
                        <table className="w-full text-left border-collapse">
                          <thead className="bg-gray-100/80">
                            <tr>
                              <th className="px-3 py-2 text-[10px] font-black uppercase tracking-widest text-gray-500">E-mail</th>
                              <th className="px-3 py-2 text-[10px] font-black uppercase tracking-widest text-gray-500 text-right">Ações</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-100">
                            {(() => {
                              const filtered = recipients.filter(e => e.toLowerCase().includes(emailSearch.toLowerCase()));
                              const paginated = filtered.slice((emailPage - 1) * itemsPerPage, emailPage * itemsPerPage);
                              
                              if (paginated.length === 0) {
                                return (
                                  <tr>
                                    <td colSpan={2} className="px-3 py-8 text-center text-xs text-gray-400 italic">Nenhum e-mail encontrado.</td>
                                  </tr>
                                );
                              }

                              return paginated.map((email) => {
                                const idx = recipients.indexOf(email);
                                return (
                                  <tr key={idx} className="group hover:bg-white transition-colors">
                                    <td className="px-3 py-2">
                                      {editingEmail?.index === idx ? (
                                        <input 
                                          type="email" 
                                          value={editingEmail.value}
                                          onChange={(e) => setEditingEmail({ ...editingEmail, value: e.target.value })}
                                          className="w-full px-2 py-1 border border-gray-300 rounded text-xs focus:outline-none"
                                          autoFocus
                                          onBlur={saveEdit}
                                          onKeyDown={(e) => e.key === 'Enter' && saveEdit()}
                                        />
                                      ) : (
                                        <span className="text-xs font-medium text-gray-700 truncate block max-w-[150px]">{email}</span>
                                      )}
                                    </td>
                                    <td className="px-3 py-2 text-right">
                                      <div className="flex items-center justify-end gap-1">
                                        <button onClick={() => startEdit(idx, email)} className="text-blue-600 hover:bg-blue-50 p-1.5 rounded-md transition-colors">
                                          <Edit2 size={12} />
                                        </button>
                                        <button onClick={() => removeRecipient(idx)} className="text-dhl-red hover:bg-red-50 p-1.5 rounded-md transition-colors">
                                          <Trash2 size={12} />
                                        </button>
                                      </div>
                                    </td>
                                  </tr>
                                );
                              });
                            })()}
                          </tbody>
                        </table>
                      </div>

                      {/* Pagination */}
                      {(() => {
                        const filtered = recipients.filter(e => e.toLowerCase().includes(emailSearch.toLowerCase()));
                        const totalPages = Math.ceil(filtered.length / itemsPerPage);
                        if (totalPages <= 1) return null;
                        return (
                          <div className="flex items-center justify-between pt-2">
                            <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pág. {emailPage} de {totalPages}</span>
                            <div className="flex gap-1">
                              <button 
                                disabled={emailPage === 1}
                                onClick={() => setEmailPage(p => p - 1)}
                                className="p-1 rounded border border-gray-200 disabled:opacity-30 hover:bg-gray-100"
                              >
                                <ChevronLeft size={14} />
                              </button>
                              <button 
                                disabled={emailPage === totalPages}
                                onClick={() => setEmailPage(p => p + 1)}
                                className="p-1 rounded border border-gray-200 disabled:opacity-30 hover:bg-gray-100"
                              >
                                <ChevronRight size={14} />
                              </button>
                            </div>
                          </div>
                        );
                      })()}
                    </div>

                    {/* Mandatory Tags Section */}
                    <div className="space-y-4 flex flex-col h-full">
                      <div className="flex items-center justify-between">
                        <h4 className="text-sm font-black uppercase tracking-widest text-gray-400 flex items-center gap-2">
                          <Braces size={16} /> Campos Obrigatórios
                        </h4>
                        <div className="relative">
                          <Search size={14} className="absolute left-2 top-1/2 -translate-y-1/2 text-gray-400" />
                          <input 
                            type="text"
                            placeholder="Filtrar..."
                            value={tagSearch}
                            onChange={(e) => { setTagSearch(e.target.value); setTagPage(1); }}
                            className="pl-7 pr-2 py-1 border border-gray-200 rounded text-[10px] focus:outline-none focus:ring-1 focus:ring-dhl-red/20 w-32"
                          />
                        </div>
                      </div>

                      <form onSubmit={addTag} className="flex flex-col gap-2">
                        <input 
                          type="text" 
                          value={newTagName}
                          onChange={(e) => setNewTagName(e.target.value)}
                          placeholder="Nome (ex: Número da Nota)..."
                          className="px-3 py-2 border border-gray-300 rounded-md text-sm focus:outline-none focus:ring-2 focus:ring-dhl-red/20"
                        />
                        <div className="flex gap-2">
                          <input 
                            type="text" 
                            value={newTagRef}
                            onChange={(e) => setNewTagRef(e.target.value)}
                            placeholder="Ref. XML (ex: nNF)..."
                            className="flex-1 px-3 py-2 border border-gray-300 rounded-md text-sm font-mono focus:outline-none focus:ring-2 focus:ring-dhl-red/20"
                          />
                          <button type="submit" className="bg-dhl-dark text-white p-2 rounded-md hover:bg-black">
                            <Plus size={16} />
                          </button>
                        </div>
                      </form>

                      <div className="flex-1 border border-gray-100 rounded-xl overflow-hidden bg-gray-50/50">
                        <table className="w-full text-left border-collapse">
                          <thead className="bg-gray-100/80">
                            <tr>
                              <th className="px-3 py-2 text-[10px] font-black uppercase tracking-widest text-gray-500">Campo</th>
                              <th className="px-3 py-2 text-[10px] font-black uppercase tracking-widest text-gray-500 text-right">Ações</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-100">
                            {(() => {
                              const filtered = mandatoryTags.filter(t => 
                                t.name.toLowerCase().includes(tagSearch.toLowerCase()) || 
                                t.tag.toLowerCase().includes(tagSearch.toLowerCase())
                              );
                              const paginated = filtered.slice((tagPage - 1) * itemsPerPage, tagPage * itemsPerPage);
                              
                              if (paginated.length === 0) {
                                return (
                                  <tr>
                                    <td colSpan={2} className="px-3 py-8 text-center text-xs text-gray-400 italic">Nenhum campo encontrado.</td>
                                  </tr>
                                );
                              }

                              return paginated.map((m) => (
                                <tr key={m.tag} className="group hover:bg-white transition-colors">
                                  <td className="px-3 py-2">
                                    <div className="flex flex-col">
                                      <span className="text-xs font-bold text-gray-700 truncate block max-w-[120px]">{m.name}</span>
                                      <span className="text-[9px] font-mono text-gray-400 uppercase">{m.tag}</span>
                                    </div>
                                  </td>
                                  <td className="px-3 py-2 text-right">
                                    <button onClick={() => removeTag(m.tag)} className="text-dhl-red hover:bg-red-50 p-1.5 rounded-md transition-colors">
                                      <Trash2 size={12} />
                                    </button>
                                  </td>
                                </tr>
                              ));
                            })()}
                          </tbody>
                        </table>
                      </div>

                      {/* Pagination */}
                      {(() => {
                        const filtered = mandatoryTags.filter(t => 
                          t.name.toLowerCase().includes(tagSearch.toLowerCase()) || 
                          t.tag.toLowerCase().includes(tagSearch.toLowerCase())
                        );
                        const totalPages = Math.ceil(filtered.length / itemsPerPage);
                        if (totalPages <= 1) return null;
                        return (
                          <div className="flex items-center justify-between pt-2">
                            <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pág. {tagPage} de {totalPages}</span>
                            <div className="flex gap-1">
                              <button 
                                disabled={tagPage === 1}
                                onClick={() => setTagPage(p => p - 1)}
                                className="p-1 rounded border border-gray-200 disabled:opacity-30 hover:bg-gray-100"
                              >
                                <ChevronLeft size={14} />
                              </button>
                              <button 
                                disabled={tagPage === totalPages}
                                onClick={() => setTagPage(p => p + 1)}
                                className="p-1 rounded border border-gray-200 disabled:opacity-30 hover:bg-gray-100"
                              >
                                <ChevronRight size={14} />
                              </button>
                            </div>
                          </div>
                        );
                      })()}
                    </div>

                    {/* OS Rules Section */}
                    <div className="space-y-4 flex flex-col h-full">
                      <div className="flex items-center justify-between">
                        <h4 className="text-sm font-black uppercase tracking-widest text-gray-400 flex items-center gap-2">
                          <ShieldAlert size={16} /> Regras de OS (Regex)
                        </h4>
                        <div className="relative">
                          <Search size={14} className="absolute left-2 top-1/2 -translate-y-1/2 text-gray-400" />
                          <input 
                            type="text"
                            placeholder="Filtrar..."
                            value={patternSearch}
                            onChange={(e) => { setPatternSearch(e.target.value); setPatternPage(1); }}
                            className="pl-7 pr-2 py-1 border border-gray-200 rounded text-[10px] focus:outline-none focus:ring-1 focus:ring-dhl-red/20 w-32"
                          />
                        </div>
                      </div>

                      <form onSubmit={addPattern} className="flex gap-2">
                        <input 
                          type="text" 
                          value={newPattern}
                          onChange={(e) => setNewPattern(e.target.value)}
                          placeholder="Regex (ex: OS:\s+)..."
                          className="flex-1 px-3 py-2 border border-gray-300 rounded-md text-sm font-mono focus:outline-none focus:ring-2 focus:ring-dhl-red/20"
                        />
                        <button type="submit" className="bg-dhl-dark text-white p-2 rounded-md hover:bg-black">
                          <Plus size={16} />
                        </button>
                      </form>

                      <div className="flex-1 border border-gray-100 rounded-xl overflow-hidden bg-gray-50/50">
                        <table className="w-full text-left border-collapse">
                          <thead className="bg-gray-100/80">
                            <tr>
                              <th className="px-3 py-2 text-[10px] font-black uppercase tracking-widest text-gray-500">Padrão Regex</th>
                              <th className="px-3 py-2 text-[10px] font-black uppercase tracking-widest text-gray-500 text-right">Ações</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-100">
                            {(() => {
                              const filtered = osForbiddenPatterns.filter(p => p.toLowerCase().includes(patternSearch.toLowerCase()));
                              const paginated = filtered.slice((patternPage - 1) * itemsPerPage, patternPage * itemsPerPage);
                              
                              if (paginated.length === 0) {
                                return (
                                  <tr>
                                    <td colSpan={2} className="px-3 py-8 text-center text-xs text-gray-400 italic">Nenhuma regra encontrada.</td>
                                  </tr>
                                );
                              }

                              return paginated.map((pattern) => (
                                <tr key={pattern} className="group hover:bg-white transition-colors">
                                  <td className="px-3 py-2">
                                    <span className="text-[10px] font-mono text-gray-600 truncate block max-w-[150px]">{pattern}</span>
                                  </td>
                                  <td className="px-3 py-2 text-right">
                                    <button onClick={() => removePattern(pattern)} className="text-dhl-red hover:bg-red-50 p-1.5 rounded-md transition-colors">
                                      <Trash2 size={12} />
                                    </button>
                                  </td>
                                </tr>
                              ));
                            })()}
                          </tbody>
                        </table>
                      </div>

                      {/* Pagination */}
                      {(() => {
                        const filtered = osForbiddenPatterns.filter(p => p.toLowerCase().includes(patternSearch.toLowerCase()));
                        const totalPages = Math.ceil(filtered.length / itemsPerPage);
                        if (totalPages <= 1) return null;
                        return (
                          <div className="flex items-center justify-between pt-2">
                            <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pág. {patternPage} de {totalPages}</span>
                            <div className="flex gap-1">
                              <button 
                                disabled={patternPage === 1}
                                onClick={() => setPatternPage(p => p - 1)}
                                className="p-1 rounded border border-gray-200 disabled:opacity-30 hover:bg-gray-100"
                              >
                                <ChevronLeft size={14} />
                              </button>
                              <button 
                                disabled={patternPage === totalPages}
                                onClick={() => setPatternPage(p => p + 1)}
                                className="p-1 rounded border border-gray-200 disabled:opacity-30 hover:bg-gray-100"
                              >
                                <ChevronRight size={14} />
                              </button>
                            </div>
                          </div>
                        );
                      })()}
                    </div>

                    {/* Registered Products Section */}
                    <div className="space-y-4 flex flex-col h-full">
                      <div className="flex items-center justify-between">
                        <h4 className="text-sm font-black uppercase tracking-widest text-gray-400 flex items-center gap-2">
                          <Search size={16} /> Produtos Cadastrados (NTV)
                        </h4>
                        <div className="relative">
                          <Search size={14} className="absolute left-2 top-1/2 -translate-y-1/2 text-gray-400" />
                          <input 
                            type="text"
                            placeholder="Filtrar..."
                            value={productSearch}
                            onChange={(e) => { setProductSearch(e.target.value); setProductPage(1); }}
                            className="pl-7 pr-2 py-1 border border-gray-200 rounded text-[10px] focus:outline-none focus:ring-1 focus:ring-dhl-red/20 w-32"
                          />
                        </div>
                      </div>

                      <form onSubmit={addProduct} className="flex gap-2">
                        <input 
                          type="text" 
                          value={newProduct}
                          onChange={(e) => setNewProduct(e.target.value)}
                          placeholder="Produto (ex: PWI)..."
                          className="flex-1 px-3 py-2 border border-gray-300 rounded-md text-sm font-bold focus:outline-none focus:ring-2 focus:ring-dhl-red/20"
                        />
                        <button type="submit" className="bg-dhl-dark text-white p-2 rounded-md hover:bg-black">
                          <Plus size={16} />
                        </button>
                      </form>

                      <div className="flex-1 border border-gray-100 rounded-xl overflow-hidden bg-gray-50/50">
                        <table className="w-full text-left border-collapse">
                          <thead className="bg-gray-100/80">
                            <tr>
                              <th className="px-3 py-2 text-[10px] font-black uppercase tracking-widest text-gray-500">Produto</th>
                              <th className="px-3 py-2 text-[10px] font-black uppercase tracking-widest text-gray-500 text-right">Ações</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-100">
                            {(() => {
                              const filtered = registeredProducts.filter(p => p.toLowerCase().includes(productSearch.toLowerCase()));
                              const paginated = filtered.slice((productPage - 1) * itemsPerPage, productPage * itemsPerPage);
                              
                              if (paginated.length === 0) {
                                return (
                                  <tr>
                                    <td colSpan={2} className="px-3 py-8 text-center text-xs text-gray-400 italic">Nenhum produto cadastrado.</td>
                                  </tr>
                                );
                              }

                              return paginated.map((product) => (
                                <tr key={product} className="group hover:bg-white transition-colors">
                                  <td className="px-3 py-2">
                                    <span className="text-xs font-bold text-gray-700 truncate block max-w-[150px]">{product}</span>
                                  </td>
                                  <td className="px-3 py-2 text-right">
                                    <button onClick={() => removeProduct(product)} className="text-dhl-red hover:bg-red-50 p-1.5 rounded-md transition-colors">
                                      <Trash2 size={12} />
                                    </button>
                                  </td>
                                </tr>
                              ));
                            })()}
                          </tbody>
                        </table>
                      </div>

                      {/* Pagination */}
                      {(() => {
                        const filtered = registeredProducts.filter(p => p.toLowerCase().includes(productSearch.toLowerCase()));
                        const totalPages = Math.ceil(filtered.length / itemsPerPage);
                        if (totalPages <= 1) return null;
                        return (
                          <div className="flex items-center justify-between pt-2">
                            <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Pág. {productPage} de {totalPages}</span>
                            <div className="flex gap-1">
                              <button 
                                disabled={productPage === 1}
                                onClick={() => setProductPage(p => p - 1)}
                                className="p-1 rounded border border-gray-200 disabled:opacity-30 hover:bg-gray-100"
                              >
                                <ChevronLeft size={14} />
                              </button>
                              <button 
                                disabled={productPage === totalPages}
                                onClick={() => setProductPage(p => p + 1)}
                                className="p-1 rounded border border-gray-200 disabled:opacity-30 hover:bg-gray-100"
                              >
                                <ChevronRight size={14} />
                              </button>
                            </div>
                          </div>
                        );
                      })()}
                    </div>
                  </div>
                </div>
              </motion.div>
            )}
          </AnimatePresence>

          <motion.div 
            onDrop={onDrop}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            className={`
              relative border-4 border-dashed rounded-2xl p-12 transition-all duration-300
              flex flex-col items-center justify-center gap-4 cursor-pointer
              ${isDragging ? 'border-dhl-red bg-dhl-yellow/20 scale-[1.01]' : 'border-dhl-yellow bg-white hover:border-dhl-red'}
            `}
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
          >
            <input 
              type="file" 
              multiple 
              accept=".xml" 
              className="absolute inset-0 opacity-0 cursor-pointer"
              onChange={(e) => e.target.files && handleFiles(e.target.files)}
            />
            <div className="bg-dhl-yellow p-6 rounded-full shadow-inner">
              <Upload className="w-12 h-12 text-dhl-red" />
            </div>
            <div className="text-center">
              <p className="text-xl font-bold">Arraste seus arquivos XML aqui</p>
              <p className="text-gray-500">ou clique para selecionar do seu computador</p>
            </div>
            <div className="flex gap-4 mt-2">
              <span className="bg-gray-100 text-[10px] font-bold px-2 py-1 rounded uppercase">Suporta múltiplos arquivos</span>
              <span className="bg-gray-100 text-[10px] font-bold px-2 py-1 rounded uppercase">Formato NFe 4.00</span>
            </div>
          </motion.div>
        </section>

        {/* Results Section */}
        <section className="space-y-6">
          {results.length > 0 && (
            <motion.div 
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className="bg-white p-4 rounded-xl shadow-sm border border-gray-100 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4"
            >
              <div className="relative">
                <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                <input 
                  type="text"
                  placeholder="nNF..."
                  value={resultsFilters.nNF}
                  onChange={(e) => setResultsFilters(prev => ({ ...prev, nNF: e.target.value }))}
                  className="w-full pl-9 pr-3 py-2 text-xs border border-gray-200 rounded-lg focus:ring-2 focus:ring-dhl-red focus:border-transparent outline-none"
                />
              </div>
              <div className="relative">
                <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                <input 
                  type="text"
                  placeholder="CNPJ..."
                  value={resultsFilters.cnpj}
                  onChange={(e) => setResultsFilters(prev => ({ ...prev, cnpj: e.target.value }))}
                  className="w-full pl-9 pr-3 py-2 text-xs border border-gray-200 rounded-lg focus:ring-2 focus:ring-dhl-red focus:border-transparent outline-none"
                />
              </div>
              <div className="relative">
                <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                <input 
                  type="text"
                  placeholder="NCM..."
                  value={resultsFilters.ncm}
                  onChange={(e) => setResultsFilters(prev => ({ ...prev, ncm: e.target.value }))}
                  className="w-full pl-9 pr-3 py-2 text-xs border border-gray-200 rounded-lg focus:ring-2 focus:ring-dhl-red focus:border-transparent outline-none"
                />
              </div>
              <div className="relative">
                <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                <input 
                  type="text"
                  placeholder="OS..."
                  value={resultsFilters.os}
                  onChange={(e) => setResultsFilters(prev => ({ ...prev, os: e.target.value }))}
                  className="w-full pl-9 pr-3 py-2 text-xs border border-gray-200 rounded-lg focus:ring-2 focus:ring-dhl-red focus:border-transparent outline-none"
                />
              </div>
              <div className="relative">
                <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                <input 
                  type="text"
                  placeholder="Produto..."
                  value={resultsFilters.xProd}
                  onChange={(e) => setResultsFilters(prev => ({ ...prev, xProd: e.target.value }))}
                  className="w-full pl-9 pr-3 py-2 text-xs border border-gray-200 rounded-lg focus:ring-2 focus:ring-dhl-red focus:border-transparent outline-none"
                />
              </div>
              {Object.values(resultsFilters).some(v => v !== '') && (
                <div className="lg:col-span-5 flex justify-end">
                  <button 
                    onClick={() => setResultsFilters({ nNF: '', cnpj: '', ncm: '', os: '', xProd: '' })}
                    className="flex items-center gap-2 text-[10px] font-black uppercase tracking-widest text-dhl-red hover:text-red-700 transition-colors"
                  >
                    <X size={14} /> Limpar Filtros
                  </button>
                </div>
              )}
            </motion.div>
          )}

          <AnimatePresence mode="popLayout">
            {(() => {
              const filtered = results.filter(r => {
                const nNFMatch = r.nNF.toLowerCase().includes(resultsFilters.nNF.toLowerCase());
                const cnpjMatch = r.cnpj.toLowerCase().includes(resultsFilters.cnpj.toLowerCase());
                const ncmMatch = r.ncm.toLowerCase().includes(resultsFilters.ncm.toLowerCase());
                const osMatch = r.osField.toLowerCase().includes(resultsFilters.os.toLowerCase());
                const xProdMatch = r.xProd.toLowerCase().includes(resultsFilters.xProd.toLowerCase());
                return nNFMatch && cnpjMatch && ncmMatch && osMatch && xProdMatch;
              });

              if (results.length > 0 && filtered.length === 0) {
                return (
                  <motion.div 
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    className="text-center py-12 bg-white rounded-xl border border-dashed border-gray-200"
                  >
                    <Search size={40} className="mx-auto text-gray-200 mb-4" />
                    <p className="text-gray-500 font-bold">Nenhum resultado corresponde aos filtros aplicados.</p>
                    <button 
                      onClick={() => setResultsFilters({ nNF: '', cnpj: '', ncm: '', os: '', xProd: '' })}
                      className="mt-4 text-dhl-red font-black uppercase text-xs hover:underline"
                    >
                      Limpar todos os filtros
                    </button>
                  </motion.div>
                );
              }

              return results.map((result, idx) => {
                const nNFMatch = result.nNF.toLowerCase().includes(resultsFilters.nNF.toLowerCase());
                const cnpjMatch = result.cnpj.toLowerCase().includes(resultsFilters.cnpj.toLowerCase());
                const ncmMatch = result.ncm.toLowerCase().includes(resultsFilters.ncm.toLowerCase());
                const osMatch = result.osField.toLowerCase().includes(resultsFilters.os.toLowerCase());
                const xProdMatch = result.xProd.toLowerCase().includes(resultsFilters.xProd.toLowerCase());
                
                if (!(nNFMatch && cnpjMatch && ncmMatch && osMatch && xProdMatch)) return null;

                return (
                  <motion.div
                    key={`${result.fileName}-${idx}`}
                    initial={{ opacity: 0, x: -20 }}
                    animate={{ opacity: 1, x: 0 }}
                    exit={{ opacity: 0, scale: 0.95 }}
                    className="glass-card rounded-xl overflow-hidden shadow-sm border-l-8 border-l-dhl-yellow"
                  >
                <div className="p-5 bg-white flex items-center justify-between border-b border-gray-100">
                  <div className="flex items-center gap-4">
                    <div className={`p-2 rounded-lg ${result.isValid ? 'bg-green-100 text-green-600' : 'bg-red-100 text-red-600'}`}>
                      {result.isValid ? <CheckCircle2 size={24} /> : <XCircle size={24} />}
                    </div>
                    <div>
                      <h3 className="font-bold text-lg flex items-center gap-2">
                        {result.fileName}
                        {result.isValid ? (
                          <span className="text-[10px] bg-green-500 text-white px-2 py-0.5 rounded-full uppercase tracking-tighter">Válido</span>
                        ) : (
                          <span className="text-[10px] bg-red-500 text-white px-2 py-0.5 rounded-full uppercase tracking-tighter">Incompleto</span>
                        )}
                      </h3>
                      <p className="text-xs text-gray-400 font-mono uppercase">Hash: {Math.random().toString(36).substring(7).toUpperCase()}</p>
                    </div>
                  </div>
                  <div className="flex items-center gap-2">
                    <button 
                      onClick={() => downloadXml(result)}
                      className="text-gray-300 hover:text-dhl-dark p-2 transition-colors"
                      title="Baixar XML"
                    >
                      <Download size={20} />
                    </button>
                    <button 
                      onClick={() => toggleExpand(idx)}
                      className={`p-2 rounded-lg transition-all flex items-center gap-2 font-bold text-xs uppercase tracking-widest ${expandedIndices.includes(idx) ? 'bg-dhl-yellow text-dhl-dark' : 'text-gray-400 hover:bg-gray-100'}`}
                      title={expandedIndices.includes(idx) ? "Ocultar detalhes" : "Ver todos os campos"}
                    >
                      {expandedIndices.includes(idx) ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
                    </button>
                    <button 
                      onClick={() => removeResult(idx)}
                      className="text-gray-300 hover:text-dhl-red p-2 transition-colors"
                      title="Remover"
                    >
                      <Trash2 size={20} />
                    </button>
                  </div>
                </div>

                <div className="p-6 grid grid-cols-1 lg:grid-cols-3 gap-8">
                  {/* Data Table */}
                  <div className="lg:col-span-2">
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="text-left border-b border-gray-100 italic font-serif text-gray-400">
                          <th className="pb-3 font-medium uppercase text-[11px] tracking-wider">Campo</th>
                          <th className="pb-3 font-medium uppercase text-[11px] tracking-wider">Valor Extraído</th>
                          <th className="pb-3 font-medium uppercase text-[11px] tracking-wider">Status</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-50">
                        {mandatoryTags.map(m => {
                          // Case-insensitive lookup in extractedFields
                          const actualKey = Object.keys(result.extractedFields).find(k => k.toLowerCase() === m.tag.toLowerCase());
                          const value = actualKey ? result.extractedFields[actualKey] : "";
                          
                          const label = `${m.name} (${m.tag})`;

                          // Special handling for xProd with NTV check
                          if (m.tag.toLowerCase() === 'xprod') {
                            return (
                              <tr key={m.tag} className="group hover:bg-gray-50 transition-colors">
                                <td className="py-4 font-bold text-gray-600">Validação de Item (xProd)</td>
                                <td className="py-4 font-mono">
                                  <div className="flex flex-col gap-1">
                                    <span className="text-xs">{value || "---"}</span>
                                    {value && (
                                      <div className="flex items-center gap-2">
                                        {result.ntvStatus === 'loading' ? (
                                          <span className="text-[10px] text-blue-500 flex items-center gap-1 animate-pulse">
                                            <Loader2 size={10} className="animate-spin" /> Verificando NTV...
                                          </span>
                                        ) : result.ntvStatus === 'registered' ? (
                                          <span className="text-[10px] text-green-600 font-bold flex items-center gap-1">
                                            <CheckCircle2 size={10} /> Item cadastrado no sistema NTV
                                          </span>
                                        ) : result.ntvStatus === 'not_registered' ? (
                                          <span className="text-[10px] text-orange-600 font-bold flex items-center gap-1">
                                            <AlertCircle size={10} /> Item não cadastrado no NTV
                                          </span>
                                        ) : result.ntvStatus === 'error' ? (
                                          <span className="text-[10px] text-red-500 flex items-center gap-1">
                                            <XCircle size={10} /> Erro na consulta NTV
                                          </span>
                                        ) : null}
                                        <button 
                                          onClick={() => checkNtvStatus(result.fileName, value, setResults)}
                                          className="text-[9px] underline text-gray-400 hover:text-dhl-red uppercase tracking-tighter"
                                        >
                                          Revalidar
                                        </button>
                                      </div>
                                    )}
                                  </div>
                                </td>
                                <td className="py-4">
                                  {value ? <CheckCircle2 className="text-green-500" size={18} /> : <AlertCircle className="text-red-500" size={18} />}
                                </td>
                              </tr>
                            );
                          }

                          // Special handling for infCpl to show extracted OS and its status
                          if (m.tag.toLowerCase() === 'infcpl') {
                            return (
                              <tr key={m.tag} className="group hover:bg-gray-50 transition-colors">
                                <td className="py-4 font-bold text-gray-600">{m.name} (infCpl)</td>
                                <td className="py-4">
                                  <div className="flex flex-col gap-1">
                                    <span className={`font-mono px-2 py-1 rounded w-fit ${result.osField !== "Não encontrado" ? 'bg-dhl-yellow/20 text-dhl-dark font-bold' : 'text-red-500'}`}>
                                      {result.osField}
                                    </span>
                                    {result.osField !== "Não encontrado" && (
                                      <div className="flex items-center gap-2">
                                        {result.osStatus === 'loading' ? (
                                          <span className="text-[10px] text-blue-500 flex items-center gap-1 animate-pulse">
                                            <Loader2 size={10} className="animate-spin" /> Verificando OS...
                                          </span>
                                        ) : result.osStatus === 'received' ? (
                                          <span className="text-[10px] text-green-600 font-bold flex items-center gap-1">
                                            <CheckCircle2 size={10} /> OS está no sistema
                                          </span>
                                        ) : result.osStatus === 'not_received' ? (
                                          <span className="text-[10px] text-orange-600 font-bold flex items-center gap-1">
                                            <AlertCircle size={10} /> OS não está no sistema
                                          </span>
                                        ) : result.osStatus === 'error' ? (
                                          <span className="text-[10px] text-red-500 flex items-center gap-1">
                                            <XCircle size={10} /> Erro na consulta OS
                                          </span>
                                        ) : result.osStatus === 'not_found' ? (
                                          <span className="text-[10px] text-gray-400 italic">
                                            OS não identificada
                                          </span>
                                        ) : null}
                                        <button 
                                          onClick={() => checkOsStatus(result.fileName, result.osField, setResults)}
                                          className="text-[9px] underline text-gray-400 hover:text-dhl-red uppercase tracking-tighter"
                                        >
                                          Revalidar
                                        </button>
                                      </div>
                                    )}
                                  </div>
                                </td>
                                <td className="py-4">
                                  {result.osField !== "Não encontrado" ? <CheckCircle2 className="text-green-500" size={18} /> : <AlertCircle className="text-red-500" size={18} />}
                                </td>
                              </tr>
                            );
                          }

                          // Default row for other mandatory tags
                          return (
                            <tr key={m.tag} className="group hover:bg-gray-50 transition-colors">
                              <td className="py-4 font-bold text-gray-600">{label}</td>
                              <td className="py-4 font-mono">{value || "---"}</td>
                              <td className="py-4">
                                {value ? <CheckCircle2 className="text-green-500" size={18} /> : <AlertCircle className="text-red-500" size={18} />}
                              </td>
                            </tr>
                          );
                        })}
                        {/* Always show xProd if available and not already shown as mandatory */}
                        {result.xProd && !mandatoryTags.some(t => t.tag.toLowerCase() === 'xprod') && (
                          <tr className="group hover:bg-gray-50 transition-colors">
                            <td className="py-4 font-bold text-gray-600">Descrição do Produto (xProd)</td>
                            <td className="py-4 font-mono text-[10px] max-w-[300px] truncate" title={result.xProd}>
                              {result.xProd}
                            </td>
                            <td className="py-4">
                              <CheckCircle2 className="text-green-500" size={18} />
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>

                  {/* Error Log / Summary */}
                  <div className="bg-gray-50 rounded-xl p-5 border border-gray-100 flex flex-col justify-between">
                    <div>
                      <h4 className="font-black text-xs uppercase tracking-widest mb-4 flex items-center gap-2">
                        <AlertCircle size={14} className="text-dhl-red" /> Log de Validação
                      </h4>
                      {result.errors.length > 0 ? (
                        <div className="space-y-4">
                          <ul className="space-y-3">
                            {result.errors.map((err, i) => (
                              <li key={i} className="text-xs text-red-600 flex gap-2 items-start">
                                <span className="mt-1 block w-1.5 h-1.5 rounded-full bg-red-500 shrink-0" />
                                {err}
                              </li>
                            ))}
                          </ul>
                          
                          {result.spValidated && (
                             <p className="text-[10px] text-green-600 font-bold uppercase mb-2">VALIDADO NO SHAREPOINT</p>
                          )}

                          <button
                            onClick={() => handleSendReport(result, idx)}
                            disabled={sendingEmailIdx !== null || result.sent}
                            className="w-full py-2 bg-dhl-red text-white rounded-lg text-xs font-black uppercase tracking-widest flex items-center justify-center gap-2 hover:bg-red-700 transition-all shadow-md disabled:opacity-50 disabled:cursor-not-allowed"
                          >
                            {sendingEmailIdx === idx ? (
                              <><Loader2 size={16} className="animate-spin" /> ENVIANDO...</>
                            ) : result.sent ? (
                              <><CheckCircle2 size={16} /> RELATÓRIO ENVIADO</>
                            ) : (
                              <><Mail size={16} /> REPORTAR DIVERGÊNCIA</>
                            )}
                          </button>
                        </div>
                      ) : (
                        <div className="flex flex-col items-center justify-center py-4 text-center">
                          <CheckCircle2 size={32} className="text-green-500 mb-2" />
                          <p className="text-sm font-bold text-green-700">Tudo em ordem!</p>
                          {result.spValidated && (
                             <p className="mt-2 text-[10px] text-green-600 font-bold uppercase">VALIDADO NO SHAREPOINT</p>
                          )}
                        </div>
                      )}
                    </div>
                  </div>
                </div>

                {/* Expandable Section */}
                <AnimatePresence>
                  {expandedIndices.includes(idx) && (
                    <motion.div
                      initial={{ height: 0, opacity: 0 }}
                      animate={{ height: 'auto', opacity: 1 }}
                      exit={{ height: 0, opacity: 0 }}
                      className="overflow-hidden border-t border-gray-100 bg-gray-50/50"
                    >
                      <div className="p-8">
                        <div className="flex items-center gap-3 mb-6">
                          <div className="w-1 h-6 bg-dhl-yellow rounded-full" />
                          <h4 className="font-black text-sm uppercase tracking-widest">Estrutura Completa do XML (Campos Adicionais)</h4>
                        </div>
                        
                        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                          {result.allFields.map((field, fIdx) => (
                            <div key={fIdx} className="bg-white p-3 rounded-lg border border-gray-100 shadow-sm flex flex-col gap-1">
                              <span className="text-[10px] font-black text-gray-400 uppercase tracking-tighter">{field.key}</span>
                              <span className="text-xs font-mono break-all text-dhl-dark">{field.value}</span>
                            </div>
                          ))}
                          {result.allFields.length === 0 && (
                            <p className="text-xs text-gray-400 italic col-span-full">Nenhum campo adicional encontrado.</p>
                          )}
                        </div>
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </motion.div>
                );
              });
            })()}
          </AnimatePresence>

          {results.length === 0 && (
            <div className="text-center py-20 opacity-20">
              <FileText size={80} className="mx-auto mb-4" />
              <p className="text-2xl font-black italic uppercase tracking-tighter">Nenhum arquivo processado</p>
            </div>
          )}
        </section>
      </main>
      
      {/* SharePoint Manager Modal */}
      <AnimatePresence>
        {showSpManager && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="bg-white w-full max-w-5xl max-h-[90vh] rounded-3xl shadow-2xl overflow-hidden flex flex-col"
            >
              <div className="p-6 bg-dhl-dark text-white flex items-center justify-between">
                <div className="flex items-center gap-4">
                  <div className="bg-dhl-yellow p-3 rounded-2xl shadow-lg">
                    <FileSearch size={24} className="text-dhl-dark" />
                  </div>
                  <div>
                    <h2 className="text-2xl font-black tracking-tighter italic uppercase leading-none">Gerenciador SharePoint</h2>
                    <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mt-1">Pasta: SiteAssets/XMLs</p>
                  </div>
                </div>
                <div className="flex items-center gap-4">
                  <div className="flex items-center gap-4 bg-white/10 px-4 py-2 rounded-xl border border-white/10">
                    <div className="flex flex-col items-center">
                      <span className="text-lg font-black text-blue-400">{spStats.analyzed}</span>
                      <span className="text-[8px] font-bold uppercase tracking-widest opacity-60">Analisados</span>
                    </div>
                    <div className="w-px h-6 bg-white/10" />
                    <div className="flex flex-col items-center">
                      <span className="text-lg font-black text-orange-400">{spStats.pending}</span>
                      <span className="text-[8px] font-bold uppercase tracking-widest opacity-60">Pendentes</span>
                    </div>
                  </div>
                  <button 
                    onClick={fetchSpStats}
                    disabled={isFetchingSpStats}
                    className="p-3 bg-white/10 hover:bg-white/20 rounded-xl transition-all disabled:opacity-50"
                    title="Atualizar lista"
                  >
                    <RefreshCw size={20} className={isFetchingSpStats ? 'animate-spin' : ''} />
                  </button>
                  <button 
                    onClick={() => setShowSpManager(false)}
                    className="p-3 bg-white/10 hover:bg-dhl-red rounded-xl transition-all"
                  >
                    <X size={20} />
                  </button>
                </div>
              </div>

              <div className="p-6 border-b border-gray-100 bg-gray-50/50 flex flex-col md:flex-row gap-4 items-center justify-between">
                <div className="relative flex-1 w-full">
                  <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-gray-400" size={20} />
                  <input 
                    type="text"
                    placeholder="Filtrar por Nota, CNPJ, NCM, OS ou Produto..."
                    value={spManagerSearch}
                    onChange={(e) => { setSpManagerSearch(e.target.value); setSpManagerPage(1); }}
                    className="w-full pl-12 pr-4 py-3 bg-white border border-gray-200 rounded-2xl focus:outline-none focus:ring-4 focus:ring-dhl-red/5 transition-all font-medium"
                  />
                </div>
                
                <div className="flex items-center gap-2 w-full md:w-auto">
                  <div className="relative flex-1 md:w-40">
                    <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
                    <input 
                      type="date"
                      value={spManagerStartDate}
                      onChange={(e) => { setSpManagerStartDate(e.target.value); setSpManagerPage(1); }}
                      className="w-full pl-9 pr-3 py-2 bg-white border border-gray-200 rounded-xl text-xs font-bold focus:outline-none focus:ring-2 focus:ring-dhl-red/10"
                      title="Data inicial"
                    />
                  </div>
                  <span className="text-gray-400 font-bold">/</span>
                  <div className="relative flex-1 md:w-40">
                    <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
                    <input 
                      type="date"
                      value={spManagerEndDate}
                      onChange={(e) => { setSpManagerEndDate(e.target.value); setSpManagerPage(1); }}
                      className="w-full pl-9 pr-3 py-2 bg-white border border-gray-200 rounded-xl text-xs font-bold focus:outline-none focus:ring-2 focus:ring-dhl-red/10"
                      title="Data final"
                    />
                  </div>
                  {(spManagerStartDate || spManagerEndDate) && (
                    <button 
                      onClick={() => { setSpManagerStartDate(''); setSpManagerEndDate(''); setSpManagerPage(1); }}
                      className="p-2 text-gray-400 hover:text-dhl-red transition-colors"
                      title="Limpar datas"
                    >
                      <X size={16} />
                    </button>
                  )}
                </div>
              </div>

              <div className="flex-1 overflow-y-auto p-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {filteredSpFiles
                    .slice((spManagerPage - 1) * 12, spManagerPage * 12)
                    .map((file) => (
                      <div key={file.serverRelativeUrl} className="group bg-white border border-gray-100 rounded-2xl p-4 hover:shadow-md transition-all flex items-center justify-between gap-4">
                        <div className="flex items-center gap-4 flex-1 min-w-0">
                          <div className={`p-3 rounded-xl ${file.isValidated ? 'bg-blue-50 text-blue-600' : 'bg-orange-50 text-orange-600'}`}>
                            {file.isValidated ? <CheckCircle2 size={20} /> : <Clock size={20} />}
                          </div>
                          <div className="min-w-0 flex-1">
                            <h4 className="font-bold text-dhl-dark truncate text-sm" title={file.name}>{file.name}</h4>
                            <div className="flex flex-wrap items-center gap-2 mt-1">
                              <span className={`text-[9px] font-black uppercase tracking-widest px-2 py-0.5 rounded-full ${file.isValidated ? 'bg-blue-100 text-blue-700' : 'bg-orange-100 text-orange-700'}`}>
                                {file.isValidated ? 'Analisado' : 'Pendente'}
                              </span>
                              <span className="text-[9px] font-bold text-gray-400 bg-gray-50 px-2 py-0.5 rounded-full flex items-center gap-1">
                                <Calendar size={10} />
                                {new Date(file.timeCreated).toLocaleDateString('pt-BR')} {new Date(file.timeCreated).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}
                              </span>
                              {file.nNF && (
                                <span className="text-[9px] font-bold text-gray-500 bg-gray-100 px-2 py-0.5 rounded-full">
                                  NF: {file.nNF}
                                </span>
                              )}
                              {file.OS && (
                                <span className="text-[9px] font-bold text-gray-500 bg-gray-100 px-2 py-0.5 rounded-full">
                                  OS: {file.OS}
                                </span>
                              )}
                              {file.NCM && (
                                <span className="text-[9px] font-bold text-gray-500 bg-gray-100 px-2 py-0.5 rounded-full">
                                  NCM: {file.NCM}
                                </span>
                              )}
                              {file.xProd && (
                                <span className="text-[9px] font-bold text-gray-500 bg-gray-100 px-2 py-0.5 rounded-full truncate max-w-[150px]" title={file.xProd}>
                                  PROD: {file.xProd}
                                </span>
                              )}
                            </div>
                          </div>
                        </div>
                        
                        <div className="flex items-center gap-2">
                          <button
                            onClick={() => downloadFromSharePoint(file.serverRelativeUrl, file.name)}
                            className="p-2 bg-gray-50 hover:bg-gray-100 text-gray-500 rounded-lg transition-all border border-transparent hover:border-gray-200"
                            title="Baixar XML"
                          >
                            <Download size={16} />
                          </button>
                          {file.isValidated ? (
                            <button
                              onClick={() => handleRevertSpFile(file)}
                              className="px-3 py-2 bg-gray-50 hover:bg-orange-50 text-orange-600 rounded-lg text-[10px] font-black uppercase tracking-widest flex items-center gap-2 transition-all border border-transparent hover:border-orange-100"
                              title="Reverter validação"
                            >
                              <RotateCcw size={14} />
                              Reverter
                            </button>
                          ) : (
                            <button
                              onClick={() => validateSpFileManually(file)}
                              className="px-3 py-2 bg-dhl-dark hover:bg-black text-white rounded-lg text-[10px] font-black uppercase tracking-widest flex items-center gap-2 transition-all"
                              title="Enviar para validação"
                            >
                              <ArrowRight size={14} />
                              Validar
                            </button>
                          )}
                        </div>
                      </div>
                    ))}
                  
                  {filteredSpFiles.length === 0 && (
                    <div className="col-span-full py-20 text-center">
                      <FileSearch size={48} className="mx-auto mb-4 opacity-10 text-gray-400" />
                      <p className="font-black uppercase tracking-widest text-sm italic text-gray-400">Nenhum arquivo encontrado</p>
                    </div>
                  )}
                </div>
              </div>

              <div className="p-6 border-t border-gray-100 bg-gray-50/50 flex items-center justify-between">
                <p className="text-[10px] font-black uppercase tracking-widest text-gray-400 italic">
                  Mostrando {Math.min(filteredSpFiles.length, 12)} de {filteredSpFiles.length} arquivos
                </p>
                <div className="flex items-center gap-2">
                  <button 
                    onClick={() => setSpManagerPage(prev => Math.max(1, prev - 1))}
                    disabled={spManagerPage === 1}
                    className="p-2 rounded-lg hover:bg-gray-200 disabled:opacity-30 transition-all"
                  >
                    <ChevronLeft size={20} />
                  </button>
                  <span className="text-xs font-black text-dhl-dark bg-white px-4 py-2 rounded-xl shadow-sm border border-gray-200">
                    Página {spManagerPage}
                  </span>
                  <button 
                    onClick={() => setSpManagerPage(prev => prev + 1)}
                    disabled={spManagerPage * 12 >= filteredSpFiles.length}
                    className="p-2 rounded-lg hover:bg-gray-200 disabled:opacity-30 transition-all"
                  >
                    <ChevronRight size={20} />
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Revalidation Modal (formerly History) */}
      <AnimatePresence>
        {showRevalidation && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="bg-white rounded-3xl shadow-2xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col"
            >
              <div className="p-6 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
                <div className="flex items-center gap-4">
                  <div className="bg-orange-500 p-3 rounded-2xl shadow-lg">
                    <RotateCcw className="text-white" size={24} />
                  </div>
                  <div>
                    <h3 className="text-xl font-black text-dhl-dark italic uppercase tracking-tighter leading-none">
                      Revalidação de Arquivos
                    </h3>
                    <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest mt-1">
                      Arquivos marcados como validados no SharePoint
                    </p>
                  </div>
                </div>
                <button 
                  onClick={() => setShowRevalidation(false)}
                  className="p-2 hover:bg-gray-200 rounded-full transition-colors text-gray-400"
                >
                  <X size={24} />
                </button>
              </div>

              <div className="p-6 bg-white border-b border-gray-100 flex flex-col md:flex-row items-end gap-4">
                <div className="flex flex-col gap-1 w-full md:w-auto">
                  <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-1">Início</label>
                  <div className="relative">
                    <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                    <input 
                      type="date"
                      value={revalidationStartDate}
                      onChange={(e) => setRevalidationStartDate(e.target.value)}
                      className="pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-dhl-red/20 transition-all font-medium text-sm"
                    />
                  </div>
                </div>
                <div className="flex flex-col gap-1 w-full md:w-auto">
                  <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-1">Fim</label>
                  <div className="relative">
                    <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                    <input 
                      type="date"
                      value={revalidationEndDate}
                      onChange={(e) => setRevalidationEndDate(e.target.value)}
                      className="pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-dhl-red/20 transition-all font-medium text-sm"
                    />
                  </div>
                </div>
                <div className="relative flex-1 w-full">
                  <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-1">Pesquisar nos resultados</label>
                  <div className="relative mt-1">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
                    <input 
                      type="text"
                      placeholder="Buscar por arquivo, NF ou CNPJ..."
                      value={revalidationSearch}
                      onChange={(e) => { setRevalidationSearch(e.target.value); setRevalidationPage(1); }}
                      className="w-full pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-dhl-red/20 transition-all font-medium"
                    />
                  </div>
                </div>
                <button 
                  onClick={loadRevalidationFromSharePoint}
                  disabled={isFetchingRevalidation}
                  className="px-6 py-2.5 bg-dhl-dark text-white hover:bg-black rounded-xl transition-all flex items-center gap-2 font-black text-xs uppercase tracking-widest disabled:opacity-50 shadow-lg"
                >
                  <RotateCcw size={16} className={isFetchingRevalidation ? 'animate-spin' : ''} />
                  Atualizar
                </button>
              </div>

              <div className="flex-1 overflow-y-auto p-6">
                {isFetchingRevalidation && revalidationItems.length === 0 ? (
                  <div className="flex flex-col items-center justify-center py-20 text-gray-400">
                    <Loader2 size={48} className="animate-spin mb-4 opacity-20" />
                    <p className="font-black uppercase tracking-widest text-sm italic">Carregando itens...</p>
                  </div>
                ) : revalidationItems.length === 0 ? (
                  <div className="flex flex-col items-center justify-center py-20 text-gray-300">
                    <RotateCcw size={64} className="mb-4 opacity-10" />
                    <p className="font-black uppercase tracking-widest text-sm italic">Nenhum arquivo para revalidação</p>
                  </div>
                ) : (
                  <div className="space-y-4">
                    {revalidationItems
                      .filter(item => 
                        item.Title.toLowerCase().includes(revalidationSearch.toLowerCase()) ||
                        item.nNF.includes(revalidationSearch) ||
                        item.CNPJ.includes(revalidationSearch)
                      )
                      .slice((revalidationPage - 1) * 10, revalidationPage * 10)
                      .map((item) => (
                        <div key={item.Id} className="group bg-white border border-gray-100 rounded-2xl p-4 hover:shadow-md transition-all flex flex-col md:row items-center justify-between gap-4">
                          <div className="flex items-center gap-4 flex-1">
                            <div className={`p-3 rounded-xl ${item.Status === 'Validado' ? 'bg-green-50 text-green-600' : 'bg-red-50 text-dhl-red'}`}>
                              {item.Status === 'Validado' ? <CheckCircle2 size={20} /> : <AlertCircle size={20} />}
                            </div>
                            <div className="min-w-0 flex-1">
                              <h4 className="font-bold text-dhl-dark truncate text-sm" title={item.Title}>{item.Title}</h4>
                              <div className="flex flex-wrap items-center gap-x-4 gap-y-1 mt-1">
                                <span className="text-[10px] text-gray-400 flex items-center gap-1 font-bold uppercase">
                                  <Clock size={10} /> {new Date(item.ValidationDate).toLocaleString('pt-BR')}
                                </span>
                                {item.nNF && (
                                  <span className="text-[10px] bg-gray-100 text-gray-600 px-1.5 py-0.5 rounded font-mono font-bold">
                                    NF: {item.nNF}
                                  </span>
                                )}
                              </div>
                            </div>
                          </div>
                          
                          <div className="flex items-center gap-2">
                            <button
                              onClick={() => downloadFromSharePoint(item.ServerRelativeUrl, item.Title)}
                              className="p-2 bg-gray-50 hover:bg-gray-100 text-gray-500 rounded-lg transition-all border border-transparent hover:border-gray-200"
                              title="Baixar XML do SharePoint"
                            >
                              <Download size={14} />
                            </button>
                            <button
                              onClick={() => handleRevertValidation(item)}
                              className="px-3 py-2 bg-gray-50 hover:bg-orange-50 text-orange-600 rounded-lg text-[10px] font-black uppercase tracking-widest flex items-center gap-2 transition-all border border-transparent hover:border-orange-100"
                              title="Reverter validação e renomear arquivo no SharePoint"
                            >
                              <RotateCcw size={14} />
                              Reverter
                            </button>
                          </div>
                        </div>
                      ))}
                  </div>
                )}
              </div>

              {revalidationItems.length > 10 && (
                <div className="p-6 border-t border-gray-100 bg-gray-50/50 flex items-center justify-between">
                  <p className="text-xs text-gray-400 font-bold uppercase tracking-widest">
                    Página {revalidationPage} de {Math.ceil(revalidationItems.length / 10)}
                  </p>
                  <div className="flex gap-2">
                    <button 
                      disabled={revalidationPage === 1}
                      onClick={() => setRevalidationPage(p => p - 1)}
                      className="p-2 bg-white border border-gray-200 rounded-lg disabled:opacity-30 hover:bg-gray-50 transition-colors"
                    >
                      <ChevronLeft size={20} />
                    </button>
                    <button 
                      disabled={revalidationPage >= Math.ceil(revalidationItems.length / 10)}
                      onClick={() => setRevalidationPage(p => p + 1)}
                      className="p-2 bg-white border border-gray-200 rounded-lg disabled:opacity-30 hover:bg-gray-50 transition-colors"
                    >
                      <ChevronRight size={20} />
                    </button>
                  </div>
                </div>
              )}
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Full History Modal */}
      <AnimatePresence>
        {showFullHistory && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="bg-white rounded-3xl shadow-2xl w-full max-w-5xl max-h-[90vh] overflow-hidden flex flex-col"
            >
              <div className="p-6 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
                <div className="flex items-center gap-4">
                  <div className="bg-dhl-dark p-3 rounded-2xl shadow-lg">
                    <History className="text-dhl-yellow" size={24} />
                  </div>
                  <div>
                    <h3 className="text-xl font-black text-dhl-dark italic uppercase tracking-tighter leading-none">
                      Histórico Completo
                    </h3>
                    <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest mt-1">
                      Log de todas as validações realizadas no sistema
                    </p>
                  </div>
                </div>
                <button 
                  onClick={() => setShowFullHistory(false)}
                  className="p-2 hover:bg-gray-200 rounded-full transition-colors text-gray-400"
                >
                  <X size={24} />
                </button>
              </div>

              <div className="p-6 bg-white border-b border-gray-100 flex flex-col md:flex-row items-end gap-4">
                <div className="flex flex-col gap-1 w-full md:w-auto">
                  <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-1">Início</label>
                  <div className="relative">
                    <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                    <input 
                      type="date"
                      value={fullHistoryStartDate}
                      onChange={(e) => setFullHistoryStartDate(e.target.value)}
                      className="pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-dhl-red/20 transition-all font-medium text-sm"
                    />
                  </div>
                </div>
                <div className="flex flex-col gap-1 w-full md:w-auto">
                  <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-1">Fim</label>
                  <div className="relative">
                    <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                    <input 
                      type="date"
                      value={fullHistoryEndDate}
                      onChange={(e) => setFullHistoryEndDate(e.target.value)}
                      className="pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-dhl-red/20 transition-all font-medium text-sm"
                    />
                  </div>
                </div>
                <div className="relative flex-1 w-full">
                  <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-1">Pesquisar nos resultados</label>
                  <div className="relative mt-1">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
                    <input 
                      type="text"
                      placeholder="Buscar por arquivo, NF, CNPJ ou e-mail..."
                      value={fullHistorySearch}
                      onChange={(e) => { setFullHistorySearch(e.target.value); setFullHistoryPage(1); }}
                      className="w-full pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-dhl-red/20 transition-all font-medium"
                    />
                  </div>
                </div>
                <button 
                  onClick={loadFullHistoryFromSharePoint}
                  disabled={isFetchingFullHistory}
                  className="px-6 py-2.5 bg-dhl-dark text-white hover:bg-black rounded-xl transition-all flex items-center gap-2 font-black text-xs uppercase tracking-widest disabled:opacity-50 shadow-lg"
                >
                  <RotateCcw size={16} className={isFetchingFullHistory ? 'animate-spin' : ''} />
                  Atualizar
                </button>
              </div>

              <div className="flex-1 overflow-x-auto">
                <table className="w-full text-left border-collapse min-w-[800px]">
                  <thead className="sticky top-0 bg-gray-50 z-10">
                    <tr className="border-b border-gray-200">
                      <th className="p-4 text-[10px] font-black uppercase tracking-widest text-gray-400">Data/Hora</th>
                      <th className="p-4 text-[10px] font-black uppercase tracking-widest text-gray-400">Arquivo</th>
                      <th className="p-4 text-[10px] font-black uppercase tracking-widest text-gray-400">Status</th>
                      <th className="p-4 text-[10px] font-black uppercase tracking-widest text-gray-400">NF / CNPJ</th>
                      <th className="p-4 text-[10px] font-black uppercase tracking-widest text-gray-400">Usuário</th>
                      <th className="p-4 text-[10px] font-black uppercase tracking-widest text-gray-400">Origem</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {isFetchingFullHistory && fullHistory.length === 0 ? (
                      <tr>
                        <td colSpan={6} className="p-20 text-center">
                          <Loader2 size={48} className="animate-spin mx-auto mb-4 opacity-20" />
                          <p className="font-black uppercase tracking-widest text-sm italic text-gray-400">Carregando histórico...</p>
                        </td>
                      </tr>
                    ) : fullHistory.length === 0 ? (
                      <tr>
                        <td colSpan={6} className="p-20 text-center">
                          <History size={64} className="mx-auto mb-4 opacity-10 text-gray-300" />
                          <p className="font-black uppercase tracking-widest text-sm italic text-gray-300">Nenhum registro encontrado</p>
                        </td>
                      </tr>
                    ) : (
                      fullHistory
                        .filter(item => 
                          item.Title.toLowerCase().includes(fullHistorySearch.toLowerCase()) ||
                          item.nNF.includes(fullHistorySearch) ||
                          item.CNPJ.includes(fullHistorySearch) ||
                          item.UserEmail.toLowerCase().includes(fullHistorySearch.toLowerCase())
                        )
                        .slice((fullHistoryPage - 1) * 15, fullHistoryPage * 15)
                        .map((item) => (
                          <tr key={item.Id} className="hover:bg-gray-50 transition-colors">
                            <td className="p-4 text-xs font-medium text-gray-500">
                              {new Date(item.ValidationDate).toLocaleString('pt-BR')}
                            </td>
                            <td className="p-4">
                              <p className="text-xs font-bold text-dhl-dark truncate max-w-[200px]" title={item.Title}>{item.Title}</p>
                            </td>
                            <td className="p-4">
                              <span className={`text-[9px] font-black uppercase tracking-widest px-2 py-1 rounded-full ${item.Status === 'Válido' ? 'bg-green-100 text-green-700' : 'bg-red-100 text-dhl-red'}`}>
                                {item.Status}
                              </span>
                            </td>
                            <td className="p-4">
                              <div className="flex flex-col gap-0.5">
                                <span className="text-[10px] font-mono font-bold text-gray-600">NF: {item.nNF || '---'}</span>
                                <span className="text-[9px] font-mono text-gray-400">CNPJ: {item.CNPJ || '---'}</span>
                              </div>
                            </td>
                            <td className="p-4">
                              <p className="text-xs font-medium text-gray-600">{item.UserEmail}</p>
                            </td>
                            <td className="p-4">
                              <span className={`text-[9px] font-bold uppercase px-2 py-1 rounded border ${item.Source === 'SharePoint' ? 'border-blue-200 bg-blue-50 text-blue-600' : 'border-gray-200 bg-gray-50 text-gray-500'}`}>
                                {item.Source}
                              </span>
                            </td>
                          </tr>
                        ))
                    )}
                  </tbody>
                </table>
              </div>

              {fullHistory.length > 15 && (
                <div className="p-6 border-t border-gray-100 bg-gray-50/50 flex items-center justify-between">
                  <p className="text-xs text-gray-400 font-bold uppercase tracking-widest">
                    Página {fullHistoryPage} de {Math.ceil(fullHistory.length / 15)}
                  </p>
                  <div className="flex gap-2">
                    <button 
                      disabled={fullHistoryPage === 1}
                      onClick={() => setFullHistoryPage(p => p - 1)}
                      className="p-2 bg-white border border-gray-200 rounded-lg disabled:opacity-30 hover:bg-gray-50 transition-colors"
                    >
                      <ChevronLeft size={20} />
                    </button>
                    <button 
                      disabled={fullHistoryPage >= Math.ceil(fullHistory.length / 15)}
                      onClick={() => setFullHistoryPage(p => p + 1)}
                      className="p-2 bg-white border border-gray-200 rounded-lg disabled:opacity-30 hover:bg-gray-50 transition-colors"
                    >
                      <ChevronRight size={20} />
                    </button>
                  </div>
                </div>
              )}
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Footer */}
      <footer className="bg-dhl-dark text-white py-12 mt-20">
        <div className="max-w-6xl mx-auto px-6 grid grid-cols-1 md:grid-cols-3 gap-12">
          <div>
            <h3 className="text-dhl-yellow font-black text-xl italic mb-4">DHL VALIDATOR</h3>
            <p className="text-sm text-gray-400 leading-relaxed">
              Ferramenta interna para verificação rápida de conformidade de Notas Fiscais Eletrônicas. 
              Garantindo agilidade e precisão na logística de dados.
            </p>
          </div>
          <div>
            <h4 className="font-bold text-sm uppercase tracking-widest mb-4">Regras de Negócio</h4>
            <ul className="text-xs text-gray-500 space-y-2">
              <li>• Validação de nNF, CNPJ e NCM</li>
              <li>• Verificação de padrão OS (OS:00000000)</li>
              <li>• Suporte a NFe Layout 4.00</li>
              <li>• Processamento em lote</li>
            </ul>
          </div>
          <div>
            <h4 className="font-bold text-sm uppercase tracking-widest mb-4">Suporte Técnico</h4>
            <p className="text-xs text-gray-500">
              Em caso de erros no processamento, contate o departamento de TI Logística.
            </p>
            <div className="mt-4 flex gap-4">
              <div className="w-8 h-8 bg-white/10 rounded-full flex items-center justify-center">
                <FileText size={14} />
              </div>
              <div className="w-8 h-8 bg-white/10 rounded-full flex items-center justify-center">
                <AlertCircle size={14} />
              </div>
            </div>
          </div>
        </div>
        <div className="max-w-6xl mx-auto px-6 mt-12 pt-8 border-t border-white/5 flex flex-col md:row items-center justify-between gap-4">
          <p className="text-[10px] text-gray-600 uppercase tracking-widest">© 2026 DHL Logistics - Todos os direitos reservados</p>
          <div className="flex gap-6 text-[10px] text-gray-600 uppercase font-bold">
            <a href="#" className="hover:text-dhl-yellow">Privacidade</a>
            <a href="#" className="hover:text-dhl-yellow">Termos de Uso</a>
            <a href="#" className="hover:text-dhl-yellow">Cookies</a>
          </div>
        </div>
      </footer>
    </div>
  );
}
