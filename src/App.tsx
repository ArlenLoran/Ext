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
  X,
  Braces,
  ShieldAlert,
  Search,
  ChevronLeft,
  ChevronRight,
  History,
  RotateCcw,
  Clock
} from 'lucide-react';
import { sendEmail, buildXmlDivergenceEmailHtml, buildBatchXmlDivergenceEmailHtml } from './services/emailService';
import { listXmlFilesFromFolder, renameXmlFileAsValidated, revertXmlFileValidation } from './services/sharepointService';
import { SharePointListsService } from './services/sharepointLists';

interface ValidationResult {
  fileName: string;
  nNF: string;
  cnpj: string;
  ncm: string;
  osField: string;
  isValid: boolean;
  errors: string[];
  rawContent: string;
  extractedFields: Record<string, string>;
  allFields: { key: string; value: string }[];
  originalFile: File;
  sent: boolean;
  ntvStatus?: 'loading' | 'registered' | 'not_registered' | 'error';
  sharepointUrl?: string;
  spValidated?: boolean;
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

  // Pagination and Filtering State for Settings
  const [emailSearch, setEmailSearch] = useState('');
  const [emailPage, setEmailPage] = useState(1);
  const [tagSearch, setTagSearch] = useState('');
  const [tagPage, setTagPage] = useState(1);
  const [patternSearch, setPatternSearch] = useState('');
  const [patternPage, setPatternPage] = useState(1);
  const itemsPerPage = 5;

  // SharePoint Integration State
  const [isSpAvailable, setIsSpAvailable] = useState(false);
  const [isSpInitialized, setIsSpInitialized] = useState(false);
  const [isInitializingSp, setIsInitializingSp] = useState(false);

  // History State
  const [history, setHistory] = useState<any[]>([]);
  const [showHistory, setShowHistory] = useState(false);
  const [isFetchingHistory, setIsFetchingHistory] = useState(false);
  const [historySearch, setHistorySearch] = useState('');
  const [historyPage, setHistoryPage] = useState(1);

  // Validation Rules State
  const [mandatoryTags, setMandatoryTags] = useState<{ name: string, tag: string }[]>(() => {
    const saved = localStorage.getItem('dhl_mandatory_tags');
    if (saved) {
      const parsed = JSON.parse(saved);
      // Migration: if it's still an array of strings, convert to objects
      if (parsed.length > 0 && typeof parsed[0] === 'string') {
        return parsed.map((tag: string) => {
          const name = {
            nNF: "Número da Nota",
            CNPJ: "CNPJ Emitente",
            NCM: "NCM Produto",
            infCpl: "Campo OS",
            natOp: "Natureza da Operação"
          }[tag] || tag;
          return { name, tag };
        });
      }
      return parsed;
    }
    return [
      { name: "Número da Nota", tag: "nNF" },
      { name: "CNPJ Emitente", tag: "CNPJ" },
      { name: "NCM Produto", tag: "NCM" },
      { name: "Campo OS", tag: "infCpl" }
    ];
  });
  const [newTagName, setNewTagName] = useState('');
  const [newTagRef, setNewTagRef] = useState('');

  const [osForbiddenPatterns, setOsForbiddenPatterns] = useState<string[]>(() => {
    const saved = localStorage.getItem('dhl_os_forbidden_patterns');
    return saved ? JSON.parse(saved) : ["OS:\\s+\\d+", "OS:\\d+[\\.,]\\d+"];
  });
  const [newPattern, setNewPattern] = useState('');

  // Check SharePoint Context on Mount
  React.useEffect(() => {
    const available = SharePointListsService.isContextAvailable();
    setIsSpAvailable(available);
    if (available) {
      checkSpInitialization();
    }
  }, []);

  const checkSpInitialization = async () => {
    try {
      const recExists = await SharePointListsService.listExists('DHL_Recipients');
      const tagExists = await SharePointListsService.listExists('DHL_MandatoryTags');
      const patExists = await SharePointListsService.listExists('DHL_OSPatterns');
      const histExists = await SharePointListsService.listExists('DHL_ValidationHistory');
      
      if (recExists && tagExists && patExists && histExists) {
        setIsSpInitialized(true);
        loadDataFromSharePoint();
        loadHistoryFromSharePoint();
      }
    } catch (error) {
      console.error('Erro ao verificar inicialização do SharePoint:', error);
    }
  };

  const loadHistoryFromSharePoint = async () => {
    if (!SharePointListsService.isContextAvailable()) return;
    setIsFetchingHistory(true);
    try {
      const items = await SharePointListsService.getItems('DHL_ValidationHistory', {
        orderBy: 'Id desc',
        top: 500
      });
      setHistory(items);
    } catch (error) {
      console.error('Erro ao carregar histórico:', error);
    } finally {
      setIsFetchingHistory(false);
    }
  };

  const loadDataFromSharePoint = async () => {
    try {
      const spRecipients = await SharePointListsService.getItems('DHL_Recipients', { select: ['Title'] });
      if (spRecipients.length > 0) {
        setRecipients(spRecipients.map(item => item.Title));
      }

      const spTags = await SharePointListsService.getItems('DHL_MandatoryTags', { select: ['Title', 'TagRef'] });
      if (spTags.length > 0) {
        setMandatoryTags(spTags.map(item => ({ name: item.Title, tag: item.TagRef })));
      }

      const spPatterns = await SharePointListsService.getItems('DHL_OSPatterns', { select: ['Title'] });
      if (spPatterns.length > 0) {
        setOsForbiddenPatterns(spPatterns.map(item => item.Title));
      }
    } catch (error) {
      console.error('Erro ao carregar dados do SharePoint:', error);
    }
  };

  const initializeSharePoint = async () => {
    if (!isSpAvailable) {
      setNotification({ type: 'error', message: 'Contexto do SharePoint não encontrado.' });
      return;
    }

    setIsInitializingSp(true);
    try {
      // Ensure Recipients List
      await SharePointListsService.ensureList('DHL_Recipients', 'Lista de e-mails para notificações', [
        { title: 'Title', type: 'Text', required: true } // Title will be the Email
      ]);

      // Ensure Mandatory Tags List
      await SharePointListsService.ensureList('DHL_MandatoryTags', 'Campos obrigatórios para validação XML', [
        { title: 'Title', type: 'Text', required: true }, // Display Name
        { title: 'TagRef', type: 'Text', required: true } // XML Tag
      ]);

      // Ensure OS Patterns List
      await SharePointListsService.ensureList('DHL_OSPatterns', 'Padrões de regex para validação de OS', [
        { title: 'Title', type: 'Text', required: true } // Regex Pattern
      ]);

      // Ensure Validation History List
      await SharePointListsService.ensureList('DHL_ValidationHistory', 'Histórico de validações de XML', [
        { title: 'Title', type: 'Text', required: true }, // Filename
        { title: 'Status', type: 'Text', required: true },
        { title: 'nNF', type: 'Text' },
        { title: 'CNPJ', type: 'Text' },
        { title: 'ServerRelativeUrl', type: 'Text' },
        { title: 'Errors', type: 'Note' },
        { title: 'ValidationDate', type: 'DateTime' }
      ]);

      setIsSpInitialized(true);
      setNotification({ type: 'success', message: 'Listas do SharePoint inicializadas com sucesso!' });
      
      // Sync current local data to SharePoint
      await syncAllToSharePoint();
      
    } catch (error) {
      console.error('Erro ao inicializar SharePoint:', error);
      setNotification({ type: 'error', message: 'Erro ao criar listas no SharePoint.' });
    } finally {
      setIsInitializingSp(false);
      setTimeout(() => setNotification(null), 3000);
    }
  };

  const syncAllToSharePoint = async () => {
    try {
      // This is a simple sync: for each local item, upsert it to SharePoint
      for (const email of recipients) {
        await SharePointListsService.upsertItem('DHL_Recipients', `Title eq '${email}'`, { Title: email });
      }
      for (const tag of mandatoryTags) {
        await SharePointListsService.upsertItem('DHL_MandatoryTags', `TagRef eq '${tag.tag}'`, { Title: tag.name, TagRef: tag.tag });
      }
      for (const pattern of osForbiddenPatterns) {
        await SharePointListsService.upsertItem('DHL_OSPatterns', `Title eq '${pattern}'`, { Title: pattern });
      }
    } catch (error) {
      console.error('Erro ao sincronizar dados com SharePoint:', error);
    }
  };

  React.useEffect(() => {
    localStorage.setItem('dhl_recipients', JSON.stringify(recipients));
    if (isSpInitialized) {
      // Sync individual changes could be complex, for now we just save to localStorage
      // In a real app, we'd update SP on each add/remove
    }
  }, [recipients, isSpInitialized]);

  React.useEffect(() => {
    localStorage.setItem('dhl_mandatory_tags', JSON.stringify(mandatoryTags));
  }, [mandatoryTags]);

  React.useEffect(() => {
    localStorage.setItem('dhl_os_forbidden_patterns', JSON.stringify(osForbiddenPatterns));
  }, [osForbiddenPatterns]);

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

  const validateXML = async (file: File): Promise<ValidationResult> => {
    const text = await file.text();
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(text, "text/xml");
    
    const errors: string[] = [];
    
    // Helper to get element text (case-insensitive)
    const getTagValue = (tagName: string) => {
      const allElements = xmlDoc.getElementsByTagName("*");
      for (let i = 0; i < allElements.length; i++) {
        if (allElements[i].tagName.toLowerCase() === tagName.toLowerCase()) {
          return allElements[i].textContent?.trim() || "";
        }
      }
      return "";
    };

    const nNF = getTagValue("nNF");
    const cnpj = getTagValue("CNPJ");
    const ncm = getTagValue("NCM");
    const infCpl = getTagValue("infCpl");

    // Dynamic Mandatory Tags Validation
    mandatoryTags.forEach(m => {
      const val = getTagValue(m.tag);
      if (!val) {
        errors.push(`Campo obrigatório '${m.name}' não encontrado ou vazio.`);
      }
    });

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
      // Dynamic Forbidden Patterns Validation
      osForbiddenPatterns.forEach(patternStr => {
        try {
          const regex = new RegExp(patternStr, 'i');
          const match = infCpl.match(regex);
          if (match) {
            errors.push(`Aviso: Detectado possível formato inválido próximo a '${match[0]}'. O padrão correto é 'OS:62669329'.`);
          }
        } catch (e) {
          console.error("Invalid regex pattern:", patternStr);
        }
      });
    }

    // Extract all fields
    const allFields: { key: string; value: string }[] = [];
    const extractedFields: Record<string, string> = {};
    
    const traverse = (node: Node) => {
      if (node.nodeType === 1) { // Element
        const element = node as Element;
        if (element.children.length === 0 && element.textContent?.trim()) {
          const tag = element.tagName;
          const val = element.textContent.trim();
          extractedFields[tag] = val;
          // Case-insensitive check for mandatory tags
          const isMandatory = mandatoryTags.some(t => t.tag.toLowerCase() === tag.toLowerCase());
          if (!isMandatory) {
            allFields.push({ key: tag, value: val });
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
      extractedFields,
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
      
      const newResults = await handleFiles(files, spUrlMap);
      
      // Automatically rename all imported files as "validated" in SharePoint
      // so they don't appear in the next fetch.
      // We do this in the background to not block the UI.
      spFiles.forEach(async (spFile) => {
        try {
          const result = newResults.find(r => r.fileName === spFile.name);
          const newUrl = await renameXmlFileAsValidated(spFile.serverRelativeUrl);
          
          if (isSpInitialized) {
            await SharePointListsService.createItem('DHL_ValidationHistory', {
              Title: spFile.name,
              Status: result?.isValid ? 'Validado' : 'Erro',
              nNF: result?.nNF || '',
              CNPJ: result?.cnpj || '',
              ServerRelativeUrl: newUrl,
              Errors: result?.errors.join('; ') || '',
              ValidationDate: new Date().toISOString()
            });
            loadHistoryFromSharePoint();
          }

          // Update the local state with the new URL for this file
          setResults(prev => prev.map(r => 
            r.fileName === spFile.name && r.sharepointUrl === spFile.serverRelativeUrl 
            ? { ...r, sharepointUrl: newUrl, spValidated: true } 
            : r
          ));
        } catch (err) {
          console.error(`Erro ao renomear ${spFile.name}:`, err);
        }
      });

      setNotification({ type: 'success', message: `${spFiles.length} arquivos importados e validados no SharePoint!` });
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

  const removeResult = (index: number) => {
    setResults(prev => prev.filter((_, i) => i !== index));
  };

  const handleRevertValidation = async (historyItem: any) => {
    if (!isSpAvailable) return;
    setIsFetchingHistory(true);
    try {
      // 1. Revert rename in SharePoint
      await revertXmlFileValidation(historyItem.ServerRelativeUrl);
      
      // 2. Delete from history list
      await SharePointListsService.deleteItem('DHL_ValidationHistory', historyItem.Id);
      
      setNotification({ type: 'success', message: `Validação do arquivo ${historyItem.Title} revertida com sucesso!` });
      
      // 3. Refresh history
      loadHistoryFromSharePoint();
    } catch (error) {
      console.error(error);
      setNotification({ type: 'error', message: 'Erro ao reverter validação.' });
    } finally {
      setIsFetchingHistory(false);
      setTimeout(() => setNotification(null), 3000);
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
              {/* Secondary Actions Group */}
              <div className="flex items-center bg-gray-50 border border-gray-200 rounded-xl p-1">
                <button 
                  onClick={() => setShowHistory(true)}
                  className="p-2 text-gray-500 hover:bg-gray-200 rounded-lg transition-all flex items-center gap-2 font-bold text-xs uppercase tracking-widest"
                  title="Histórico de Validações"
                >
                  <History size={16} />
                  <span className="hidden sm:inline">Histórico</span>
                </button>

                <div className="w-px h-6 bg-gray-200 mx-1" />

                <button 
                  onClick={() => setShowSettings(!showSettings)}
                  className={`p-2 rounded-lg transition-all flex items-center gap-2 font-bold text-xs uppercase tracking-widest ${showSettings ? 'bg-dhl-dark text-white' : 'text-gray-500 hover:bg-gray-200'}`}
                  title="Configurações do Sistema"
                >
                  <Settings size={16} />
                  <span className="hidden sm:inline">Configurações</span>
                </button>
                
                {results.length > 0 && (
                  <>
                    <div className="w-px h-6 bg-gray-200 mx-1" />
                    <button 
                      onClick={clearAll}
                      className="p-2 text-red-500 hover:bg-red-50 rounded-lg transition-all flex items-center gap-2 font-bold text-xs uppercase tracking-widest"
                      title="Limpar Tudo"
                    >
                      <Trash2 size={16} />
                      <span className="hidden sm:inline">Limpar</span>
                    </button>
                  </>
                )}
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
                  <div className="flex items-center gap-2">
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

                          // Special handling for NCM with NTV check
                          if (m.tag.toLowerCase() === 'ncm') {
                            return (
                              <tr key={m.tag} className="group hover:bg-gray-50 transition-colors">
                                <td className="py-4 font-bold text-gray-600">{label}</td>
                                <td className="py-4 font-mono">
                                  <div className="flex flex-col gap-1">
                                    <span>{value || "---"}</span>
                                    {value && (
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
                                          onClick={() => checkNtvStatus(idx, value)}
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

                          // Special handling for infCpl to show extracted OS
                          if (m.tag.toLowerCase() === 'infcpl') {
                            return (
                              <tr key={m.tag} className="group hover:bg-gray-50 transition-colors">
                                <td className="py-4 font-bold text-gray-600">{m.name} (infCpl)</td>
                                <td className="py-4">
                                  <span className={`font-mono px-2 py-1 rounded ${result.osField !== "Não encontrado" ? 'bg-dhl-yellow/20 text-dhl-dark font-bold' : 'text-red-500'}`}>
                                    {result.osField}
                                  </span>
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
      
      {/* History Modal */}
      <AnimatePresence>
        {showHistory && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="bg-white rounded-3xl shadow-2xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col"
            >
              <div className="p-6 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
                <div className="flex items-center gap-4">
                  <div className="bg-dhl-dark p-3 rounded-2xl shadow-lg">
                    <History className="text-dhl-yellow" size={24} />
                  </div>
                  <div>
                    <h3 className="text-xl font-black text-dhl-dark italic uppercase tracking-tighter leading-none">
                      Histórico de Validações
                    </h3>
                    <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest mt-1">
                      Registros persistidos no SharePoint
                    </p>
                  </div>
                </div>
                <button 
                  onClick={() => setShowHistory(false)}
                  className="p-2 hover:bg-gray-200 rounded-full transition-colors text-gray-400"
                >
                  <X size={24} />
                </button>
              </div>

              <div className="p-6 bg-white border-b border-gray-100 flex flex-col md:row items-center gap-4">
                <div className="relative flex-1 w-full">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
                  <input 
                    type="text"
                    placeholder="Buscar por arquivo, NF ou CNPJ..."
                    value={historySearch}
                    onChange={(e) => { setHistorySearch(e.target.value); setHistoryPage(1); }}
                    className="w-full pl-10 pr-4 py-3 bg-gray-50 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-dhl-red/20 transition-all font-medium"
                  />
                </div>
                <button 
                  onClick={loadHistoryFromSharePoint}
                  disabled={isFetchingHistory}
                  className="px-4 py-3 bg-gray-100 hover:bg-gray-200 rounded-xl text-gray-600 transition-all flex items-center gap-2 font-bold text-xs uppercase tracking-widest disabled:opacity-50"
                >
                  <RotateCcw size={16} className={isFetchingHistory ? 'animate-spin' : ''} />
                  Atualizar
                </button>
              </div>

              <div className="flex-1 overflow-y-auto p-6">
                {isFetchingHistory && history.length === 0 ? (
                  <div className="flex flex-col items-center justify-center py-20 text-gray-400">
                    <Loader2 size={48} className="animate-spin mb-4 opacity-20" />
                    <p className="font-black uppercase tracking-widest text-sm italic">Carregando histórico...</p>
                  </div>
                ) : history.length === 0 ? (
                  <div className="flex flex-col items-center justify-center py-20 text-gray-300">
                    <History size={64} className="mb-4 opacity-10" />
                    <p className="font-black uppercase tracking-widest text-sm italic">Nenhum registro encontrado</p>
                  </div>
                ) : (
                  <div className="space-y-4">
                    {history
                      .filter(item => 
                        item.Title.toLowerCase().includes(historySearch.toLowerCase()) ||
                        item.nNF.includes(historySearch) ||
                        item.CNPJ.includes(historySearch)
                      )
                      .slice((historyPage - 1) * 10, historyPage * 10)
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
                                {item.CNPJ && (
                                  <span className="text-[10px] bg-gray-100 text-gray-600 px-1.5 py-0.5 rounded font-mono font-bold">
                                    CNPJ: {item.CNPJ}
                                  </span>
                                )}
                              </div>
                            </div>
                          </div>
                          
                          <div className="flex items-center gap-2">
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

              {history.length > 10 && (
                <div className="p-6 border-t border-gray-100 bg-gray-50/50 flex items-center justify-between">
                  <p className="text-xs text-gray-400 font-bold uppercase tracking-widest">
                    Página {historyPage} de {Math.ceil(history.length / 10)}
                  </p>
                  <div className="flex gap-2">
                    <button 
                      disabled={historyPage === 1}
                      onClick={() => setHistoryPage(p => p - 1)}
                      className="p-2 bg-white border border-gray-200 rounded-lg disabled:opacity-30 hover:bg-gray-50 transition-colors"
                    >
                      <ChevronLeft size={20} />
                    </button>
                    <button 
                      disabled={historyPage >= Math.ceil(history.length / 10)}
                      onClick={() => setHistoryPage(p => p + 1)}
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
