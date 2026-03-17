/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback } from 'react';
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
  X
} from 'lucide-react';
import { sendEmail, buildXmlDivergenceEmailHtml, buildBatchXmlDivergenceEmailHtml } from './services/emailService';
import { listXmlFilesFromFolder, renameXmlFileAsValidated } from './services/sharepointService';

interface ValidationResult {
  fileName: string;
  nNF: string;
  cnpj: string;
  ncm: string;
  osField: string;
  isValid: boolean;
  errors: string[];
  rawContent: string;
  allFields: { key: string; value: string }[];
  originalFile: File;
  sent: boolean;
  ntvStatus?: 'loading' | 'registered' | 'not_registered' | 'error';
  sharepointUrl?: string;
}

export default function App() {
  const [results, setResults] = useState<ValidationResult[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [expandedIndices, setExpandedIndices] = useState<number[]>([]);
  const [sendingEmailIdx, setSendingEmailIdx] = useState<number | null>(null);
  const [isSendingBatch, setIsSendingBatch] = useState(false);
  const [isFetchingSharePoint, setIsFetchingSharePoint] = useState(false);
  const [notification, setNotification] = useState<{ type: 'success' | 'error', message: string } | null>(null);
  
  // Email management state
  const [recipients, setRecipients] = useState<string[]>(() => {
    const saved = localStorage.getItem('dhl_recipients');
    return saved ? JSON.parse(saved) : ['Arlen.Oliveira@dhl.com'];
  });
  const [newEmail, setNewEmail] = useState('');
  const [editingEmail, setEditingEmail] = useState<{ index: number, value: string } | null>(null);
  const [showSettings, setShowSettings] = useState(false);

  React.useEffect(() => {
    localStorage.setItem('dhl_recipients', JSON.stringify(recipients));
  }, [recipients]);

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
    setNewEmail('');
    setNotification({ type: 'success', message: 'E-mail adicionado com sucesso!' });
    setTimeout(() => setNotification(null), 3000);
  };

  const removeRecipient = (index: number) => {
    setRecipients(recipients.filter((_, i) => i !== index));
  };

  const startEdit = (index: number, value: string) => {
    setEditingEmail({ index, value });
  };

  const saveEdit = () => {
    if (!editingEmail) return;
    const email = editingEmail.value.trim();
    if (!email || !email.includes('@')) {
      setNotification({ type: 'error', message: 'E-mail inválido.' });
      return;
    }
    const updated = [...recipients];
    updated[editingEmail.index] = email;
    setRecipients(updated);
    setEditingEmail(null);
  };

  const toggleExpand = (index: number) => {
    setExpandedIndices(prev => 
      prev.includes(index) ? prev.filter(i => i !== index) : [...prev, index]
    );
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

  const handleConfirmSharePointValidation = async (result: ValidationResult, index: number) => {
    if (!result.sharepointUrl) return;
    
    setSendingEmailIdx(index);
    try {
      const newUrl = await renameXmlFileAsValidated(result.sharepointUrl);
      setNotification({ type: 'success', message: 'Arquivo validado e renomeado no SharePoint!' });
      
      setResults(prev => {
        const updated = [...prev];
        updated[index] = { ...updated[index], sent: true, sharepointUrl: newUrl };
        return updated;
      });
    } catch (error) {
      console.error(error);
      setNotification({ type: 'error', message: 'Falha ao renomear no SharePoint.' });
    } finally {
      setSendingEmailIdx(null);
      setTimeout(() => setNotification(null), 5000);
    }
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
      let newSpUrl = result.sharepointUrl;
      if (result.sharepointUrl) {
        try {
          newSpUrl = await renameXmlFileAsValidated(result.sharepointUrl);
        } catch (spError) {
          console.error("Erro ao renomear no SharePoint após envio:", spError);
        }
      }

      setResults(prev => {
        const updated = [...prev];
        updated[index] = { ...updated[index], sent: true, sharepointUrl: newSpUrl };
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
      const updatedResults = await Promise.all(results.map(async (r) => {
        if (r.errors.length > 0 && !r.sent) {
          let newSpUrl = r.sharepointUrl;
          if (r.sharepointUrl) {
            try {
              newSpUrl = await renameXmlFileAsValidated(r.sharepointUrl);
            } catch (spError) {
              console.error("Erro ao renomear no SharePoint (lote):", spError);
            }
          }
          return { ...r, sent: true, sharepointUrl: newSpUrl };
        }
        return r;
      }));
      
      setResults(updatedResults);
    } catch (error) {
      console.error(error);
      setNotification({ type: 'error', message: 'Falha ao enviar relatório em lote com anexos.' });
    } finally {
      setIsSendingBatch(false);
      setTimeout(() => setNotification(null), 5000);
    }
  };

  const validateXML = async (file: File): Promise<ValidationResult> => {
    const text = await file.text();
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(text, "text/xml");
    
    const errors: string[] = [];
    
    // Helper to get element text
    const getTagValue = (tagName: string) => {
      const el = xmlDoc.getElementsByTagName(tagName)[0];
      return el ? el.textContent || "" : "";
    };

    const nNF = getTagValue("nNF");
    const cnpj = getTagValue("CNPJ");
    const ncm = getTagValue("NCM");
    const infCpl = getTagValue("infCpl");

    if (!nNF) errors.push("Número da Nota (nNF) não encontrado ou vazio.");
    if (!cnpj) errors.push("CNPJ não encontrado ou vazio.");
    if (!ncm) errors.push("NCM não encontrado ou vazio.");

    // OS Validation
    const osMatch = infCpl.match(/OS:(\d+)/);
    const osValue = osMatch ? osMatch[0] : "";
    
    if (!osValue) {
      if (infCpl.toLowerCase().includes("os:")) {
        errors.push("Campo OS encontrado mas em formato inválido (deve ser 'OS:12345678' sem espaços ou pontos).");
      } else {
        errors.push("Campo OS não encontrado nas informações complementares (infCpl).");
      }
    } else {
      const malformedOS = infCpl.match(/OS:\s+\d+|OS:\d+[\.,]\d+/i);
      if (malformedOS) {
        errors.push(`Aviso: Detectado possível formato inválido próximo a '${malformedOS[0]}'. O padrão correto é 'OS:62669329'.`);
      }
    }

    // Extract all other fields
    const allFields: { key: string; value: string }[] = [];
    const mandatoryTags = ["nNF", "CNPJ", "NCM", "infCpl"];
    
    const traverse = (node: Node) => {
      if (node.nodeType === 1) { // Element
        const element = node as Element;
        if (element.children.length === 0 && element.textContent?.trim()) {
          if (!mandatoryTags.includes(element.tagName)) {
            allFields.push({ key: element.tagName, value: element.textContent.trim() });
          }
        }
        for (let i = 0; i < element.children.length; i++) {
          traverse(element.children[i]);
        }
      }
    };
    traverse(xmlDoc.documentElement);

    return {
      fileName: file.name,
      nNF,
      cnpj,
      ncm,
      osField: osValue || "Não encontrado",
      isValid: errors.length === 0,
      errors,
      rawContent: text,
      allFields,
      originalFile: file,
      sent: false
    };
  };

  const clearAll = () => setResults([]);

  const handleSharePointImport = async () => {
    setIsFetchingSharePoint(true);
    try {
      const spFiles = await listXmlFilesFromFolder('SiteAssets/XMLs');
      if (spFiles.length === 0) {
        setNotification({ type: 'error', message: 'Nenhum arquivo XML encontrado na pasta do SharePoint.' });
        return;
      }
      
      const files = spFiles.map(f => f.file);
      const spUrlMap = spFiles.reduce((acc, f) => {
        acc[f.name] = f.serverRelativeUrl;
        return acc;
      }, {} as Record<string, string>);
      
      await handleFiles(files, spUrlMap);
      setNotification({ type: 'success', message: `${spFiles.length} arquivos importados do SharePoint com sucesso!` });
    } catch (error) {
      console.error(error);
      setNotification({ type: 'error', message: error instanceof Error ? error.message : 'Erro ao importar do SharePoint.' });
    } finally {
      setIsFetchingSharePoint(false);
      setTimeout(() => setNotification(null), 5000);
    }
  };

  const checkNtvStatus = async (index: number, ncm: string) => {
    if (!ncm) return;
    
    setResults(prev => {
      const updated = [...prev];
      if (updated[index]) updated[index] = { ...updated[index], ntvStatus: 'loading' };
      return updated;
    });

    const url = "https://51a805d34213e248a3506f5db8fe28.55.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/655aac37bdea49b1b1221a2f37198754/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-2l0x4h5cwmpZ20RCIbMrzaR0860ka4aB8_dDOVQQHQ";
    
    const payload = {
      query: `SELECT * FROM PRTMST WHERE PRTNUM LIKE '%${ncm}%'`,
      id_score: "12345"
    };

    try {
      const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      const contentType = response.headers.get("content-type");
      let result;
      if (contentType && contentType.includes("application/json")) {
        result = await response.json();
      } else {
        const text = await response.text();
        try { result = JSON.parse(text); } catch { result = text; }
      }

      const isRegistered = Array.isArray(result) && result.length > 0;
      
      setResults(prev => {
        const updated = [...prev];
        if (updated[index]) {
          updated[index] = { ...updated[index], ntvStatus: isRegistered ? 'registered' : 'not_registered' };
        }
        return updated;
      });
    } catch (error) {
      console.error("Erro ao verificar NTV:", error);
      setResults(prev => {
        const updated = [...prev];
        if (updated[index]) updated[index] = { ...updated[index], ntvStatus: 'error' };
        return updated;
      });
    }
  };

  const handleFiles = async (files: FileList | File[], spUrlMap?: Record<string, string>) => {
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
    
    setResults(prev => {
      const combined = [...newResults, ...prev];
      // Trigger background NTV checks for new results
      newResults.forEach((res, i) => {
        if (res.ncm) {
          checkNtvStatus(i, res.ncm);
        }
      });
      return combined;
    });
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

  const removeResult = (index: number) => {
    setResults(prev => prev.filter((_, i) => i !== index));
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
              className={`fixed top-24 right-6 z-[100] p-4 rounded-lg shadow-2xl flex items-center gap-3 border-l-4 ${
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

        {/* Hero / Upload Section */}
        <section className="space-y-4">
          <div className="flex flex-col md:flex-row md:items-end justify-between gap-4">
            <h2 className="text-3xl font-bold tracking-tight text-dhl-dark">Validação de Notas Fiscais</h2>
            <div className="flex flex-wrap items-center gap-3">
              <button 
                onClick={() => setShowSettings(!showSettings)}
                className="bg-white text-dhl-dark border border-gray-200 px-4 py-2 rounded-md transition-all flex items-center gap-2 font-bold text-sm shadow-sm hover:bg-gray-50"
              >
                <Settings size={16} /> DESTINATÁRIOS
              </button>
              <button 
                onClick={handleSharePointImport}
                disabled={isFetchingSharePoint}
                className="bg-white text-dhl-dark border border-gray-200 px-4 py-2 rounded-md transition-all flex items-center gap-2 font-bold text-sm shadow-sm hover:bg-gray-50 disabled:opacity-50"
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
                  className="bg-dhl-red text-white px-4 py-2 rounded-md transition-all flex items-center gap-2 font-bold text-sm shadow-md hover:bg-red-700 disabled:opacity-50"
                >
                  {isSendingBatch ? (
                    <><Loader2 size={16} className="animate-spin" /> ENVIANDO LOTE...</>
                  ) : results.filter(r => r.errors.length > 0).every(r => r.sent) ? (
                    <><CheckCircle2 size={16} /> LOTE ENVIADO</>
                  ) : (
                    <><Mail size={16} /> REPORTAR TODAS AS DIVERGÊNCIAS</>
                  )}
                </button>
              )}
              {results.length > 0 && (
                <button 
                  onClick={clearAll}
                  className="text-dhl-red hover:bg-dhl-red/10 px-4 py-2 rounded-md transition-colors flex items-center gap-2 font-bold text-sm"
                >
                  <Trash2 size={16} /> LIMPAR TUDO
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
                <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-4">
                  <div className="flex items-center justify-between">
                    <h3 className="text-lg font-bold text-dhl-dark flex items-center gap-2">
                      <Mail size={20} className="text-dhl-red" /> Gerenciar Destinatários
                    </h3>
                    <button onClick={() => setShowSettings(false)} className="text-gray-400 hover:text-gray-600">
                      <X size={20} />
                    </button>
                  </div>
                  
                  <form onSubmit={addRecipient} className="flex gap-2">
                    <input 
                      type="email" 
                      value={newEmail}
                      onChange={(e) => setNewEmail(e.target.value)}
                      placeholder="Novo e-mail de destino..."
                      className="flex-1 px-4 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-dhl-red/20 focus:border-dhl-red"
                    />
                    <button 
                      type="submit"
                      className="bg-dhl-dark text-white px-4 py-2 rounded-md font-bold text-sm flex items-center gap-2 hover:bg-black transition-colors"
                    >
                      <Plus size={16} /> ADICIONAR
                    </button>
                  </form>

                  <div className="space-y-2 max-h-48 overflow-y-auto pr-2">
                    {recipients.length === 0 ? (
                      <p className="text-sm text-gray-500 italic">Nenhum e-mail cadastrado. O sistema não poderá enviar relatórios.</p>
                    ) : (
                      recipients.map((email, idx) => (
                        <div key={idx} className="flex items-center justify-between p-3 bg-gray-50 rounded-lg group">
                          {editingEmail?.index === idx ? (
                            <div className="flex-1 flex gap-2">
                              <input 
                                type="email" 
                                value={editingEmail.value}
                                onChange={(e) => setEditingEmail({ ...editingEmail, value: e.target.value })}
                                className="flex-1 px-2 py-1 border border-gray-300 rounded focus:outline-none"
                                autoFocus
                              />
                              <button onClick={saveEdit} className="text-green-600 hover:text-green-700">
                                <Save size={18} />
                              </button>
                              <button onClick={() => setEditingEmail(null)} className="text-gray-400 hover:text-gray-600">
                                <X size={18} />
                              </button>
                            </div>
                          ) : (
                            <>
                              <span className="text-sm font-medium text-gray-700">{email}</span>
                              <div className="flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                                <button 
                                  onClick={() => startEdit(idx, email)}
                                  className="text-blue-600 hover:text-blue-700 p-1"
                                >
                                  <Edit2 size={16} />
                                </button>
                                <button 
                                  onClick={() => removeRecipient(idx)}
                                  className="text-dhl-red hover:text-red-700 p-1"
                                >
                                  <Trash2 size={16} />
                                </button>
                              </div>
                            </>
                          )}
                        </div>
                      ))
                    )}
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
          <AnimatePresence mode="popLayout">
            {results.map((result, idx) => (
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
                  <button 
                    onClick={() => removeResult(idx)}
                    className="text-gray-300 hover:text-dhl-red p-2 transition-colors"
                  >
                    <Trash2 size={20} />
                  </button>
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
                        <tr className="group hover:bg-gray-50 transition-colors">
                          <td className="py-4 font-bold text-gray-600">Número da Nota (nNF)</td>
                          <td className="py-4 font-mono text-dhl-red">{result.nNF || "---"}</td>
                          <td className="py-4">
                            {result.nNF ? <CheckCircle2 className="text-green-500" size={18} /> : <AlertCircle className="text-red-500" size={18} />}
                          </td>
                        </tr>
                        <tr className="group hover:bg-gray-50 transition-colors">
                          <td className="py-4 font-bold text-gray-600">CNPJ Emitente</td>
                          <td className="py-4 font-mono">{result.cnpj || "---"}</td>
                          <td className="py-4">
                            {result.cnpj ? <CheckCircle2 className="text-green-500" size={18} /> : <AlertCircle className="text-red-500" size={18} />}
                          </td>
                        </tr>
                        <tr className="group hover:bg-gray-50 transition-colors">
                          <td className="py-4 font-bold text-gray-600">NCM Produto</td>
                          <td className="py-4 font-mono">
                            <div className="flex flex-col gap-1">
                              <span>{result.ncm || "---"}</span>
                              {result.ncm && (
                                <div className="flex items-center gap-2">
                                  {result.ntvStatus === 'loading' ? (
                                    <span className="text-[10px] text-blue-500 flex items-center gap-1 animate-pulse">
                                      <Loader2 size={10} className="animate-spin" /> Verificando NTV...
                                    </span>
                                  ) : result.ntvStatus === 'registered' ? (
                                    <span className="text-[10px] text-green-600 font-bold flex items-center gap-1">
                                      <CheckCircle2 size={10} /> Já cadastrado no sistema NTV
                                    </span>
                                  ) : result.ntvStatus === 'not_registered' ? (
                                    <span className="text-[10px] text-orange-600 font-bold flex items-center gap-1">
                                      <AlertCircle size={10} /> Não cadastrado no NTV
                                    </span>
                                  ) : result.ntvStatus === 'error' ? (
                                    <span className="text-[10px] text-red-500 flex items-center gap-1">
                                      <XCircle size={10} /> Erro na consulta NTV
                                    </span>
                                  ) : null}
                                  <button 
                                    onClick={() => checkNtvStatus(idx, result.ncm)}
                                    className="text-[9px] underline text-gray-400 hover:text-dhl-red uppercase tracking-tighter"
                                  >
                                    Revalidar
                                  </button>
                                </div>
                              )}
                            </div>
                          </td>
                          <td className="py-4">
                            {result.ncm ? <CheckCircle2 className="text-green-500" size={18} /> : <AlertCircle className="text-red-500" size={18} />}
                          </td>
                        </tr>
                        <tr className="group hover:bg-gray-50 transition-colors">
                          <td className="py-4 font-bold text-gray-600">Campo OS (infCpl)</td>
                          <td className="py-4">
                            <span className={`font-mono px-2 py-1 rounded ${result.osField !== "Não encontrado" ? 'bg-dhl-yellow/20 text-dhl-dark font-bold' : 'text-red-500'}`}>
                              {result.osField}
                            </span>
                          </td>
                          <td className="py-4">
                            {result.osField !== "Não encontrado" ? <CheckCircle2 className="text-green-500" size={18} /> : <AlertCircle className="text-red-500" size={18} />}
                          </td>
                        </tr>
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
                          {result.sharepointUrl && !result.sent && (
                            <button
                              onClick={() => handleConfirmSharePointValidation(result, idx)}
                              disabled={sendingEmailIdx !== null}
                              className="mt-4 w-full py-2 bg-green-600 text-white rounded-lg text-xs font-black uppercase tracking-widest flex items-center justify-center gap-2 hover:bg-green-700 transition-all shadow-md disabled:opacity-50"
                            >
                              {sendingEmailIdx === idx ? (
                                <><Loader2 size={16} className="animate-spin" /> PROCESSANDO...</>
                              ) : (
                                <><CheckCircle2 size={16} /> CONFIRMAR NO SHAREPOINT</>
                              )}
                            </button>
                          )}
                          {result.sharepointUrl && result.sent && (
                             <p className="mt-2 text-[10px] text-green-600 font-bold uppercase">VALIDADO NO SHAREPOINT</p>
                          )}
                        </div>
                      )}
                    </div>
                    
                    <button 
                      onClick={() => toggleExpand(idx)}
                      className="mt-6 w-full py-3 bg-white border border-gray-200 rounded-lg text-xs font-black uppercase tracking-widest flex items-center justify-center gap-2 hover:bg-dhl-yellow/10 transition-colors shadow-sm"
                    >
                      {expandedIndices.includes(idx) ? (
                        <>OCULTAR DETALHES <ChevronUp size={16} /></>
                      ) : (
                        <>VER TODOS OS CAMPOS <ChevronDown size={16} /></>
                      )}
                    </button>
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
            ))}
          </AnimatePresence>

          {results.length === 0 && (
            <div className="text-center py-20 opacity-20">
              <FileText size={80} className="mx-auto mb-4" />
              <p className="text-2xl font-black italic uppercase tracking-tighter">Nenhum arquivo processado</p>
            </div>
          )}
        </section>
      </main>

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
