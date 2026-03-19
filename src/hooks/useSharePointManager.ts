import { useState, useCallback, useEffect, useMemo } from 'react';
import { listXmlFilesFromFolder, renameXmlFileAsValidated, revertXmlFileValidation, downloadFileFromSharePoint, listAllXmlFilesFromFolder } from '../services/sharepointService';
import { SharePointListsService } from '../services/sharepointLists';
import { SpFile, SpStats, MandatoryTag } from '../types';

export function useSharePointManager(
  showNotification: (type: 'success' | 'error', message: string) => void,
  recipients: string[],
  setRecipients: (r: string[]) => void,
  mandatoryTags: MandatoryTag[],
  setMandatoryTags: (t: MandatoryTag[]) => void,
  osForbiddenPatterns: string[],
  setOsForbiddenPatterns: (p: string[]) => void
) {
  const [isSpAvailable, setIsSpAvailable] = useState(false);
  const [isSpInitialized, setIsSpInitialized] = useState(false);
  const [isInitializingSp, setIsInitializingSp] = useState(false);
  const [isFetchingSharePoint, setIsFetchingSharePoint] = useState(false);

  // Revalidation State
  const [revalidationItems, setRevalidationItems] = useState<any[]>([]);
  const [showRevalidation, setShowRevalidation] = useState(false);
  const [isFetchingRevalidation, setIsFetchingRevalidation] = useState(false);
  const [revalidationSearch, setRevalidationSearch] = useState('');
  const [revalidationPage, setRevalidationPage] = useState(1);
  const [revalidationStartDate, setRevalidationStartDate] = useState('');
  const [revalidationEndDate, setRevalidationEndDate] = useState('');

  // Full History State
  const [fullHistory, setFullHistory] = useState<any[]>([]);
  const [showFullHistory, setShowFullHistory] = useState(false);
  const [isFetchingFullHistory, setIsFetchingFullHistory] = useState(false);
  const [fullHistorySearch, setFullHistorySearch] = useState('');
  const [fullHistoryPage, setFullHistoryPage] = useState(1);
  const [fullHistoryStartDate, setFullHistoryStartDate] = useState('');
  const [fullHistoryEndDate, setFullHistoryEndDate] = useState('');

  // SharePoint Stats State
  const [spStats, setSpStats] = useState({ analyzed: 0, pending: 0 });
  const [spFilesList, setSpFilesList] = useState<SpFile[]>([]);
  const [isFetchingSpStats, setIsFetchingSpStats] = useState(false);
  const [showSpManager, setShowSpManager] = useState(false);
  const [spManagerSearch, setSpManagerSearch] = useState('');
  const [spManagerPage, setSpManagerPage] = useState(1);
  const [spManagerStartDate, setSpManagerStartDate] = useState('');
  const [spManagerEndDate, setSpManagerEndDate] = useState('');

  const filteredSpFiles = useMemo(() => {
    return spFilesList.filter(file => {
      if (spManagerStartDate) {
        const fileDate = new Date(file.timeCreated);
        const startDate = new Date(spManagerStartDate);
        startDate.setHours(0, 0, 0, 0);
        if (fileDate < startDate) return false;
      }
      
      if (spManagerEndDate) {
        const fileDate = new Date(file.timeCreated);
        const endDate = new Date(spManagerEndDate);
        endDate.setHours(23, 59, 59, 999);
        if (fileDate > endDate) return false;
      }

      const search = spManagerSearch.toLowerCase();
      if (!search) return true;
      return (
        file.name.toLowerCase().includes(search) ||
        file.nNF?.toLowerCase().includes(search) ||
        file.CNPJ?.toLowerCase().includes(search) ||
        file.OS?.toLowerCase().includes(search) ||
        file.NCM?.toLowerCase().includes(search) ||
        file.xProd?.toLowerCase().includes(search)
      );
    });
  }, [spFilesList, spManagerSearch, spManagerStartDate, spManagerEndDate]);

  const fetchSpStats = useCallback(async () => {
    if (!SharePointListsService.isContextAvailable()) return;
    setIsFetchingSpStats(true);
    try {
      const allFiles = await listAllXmlFilesFromFolder('SiteAssets/XMLs');
      
      const analyzed = allFiles.filter(f => f.isValidated).length;
      const pending = allFiles.filter(f => !f.isValidated).length;
      
      setSpStats({ analyzed, pending });

      let enrichedFiles = allFiles.map(f => ({
        name: f.name,
        serverRelativeUrl: f.serverRelativeUrl,
        isValidated: f.isValidated,
        timeCreated: f.timeCreated
      }));

      try {
        const history = await SharePointListsService.getItems('DHL_FullHistory', {
          select: ['Title', 'nNF', 'CNPJ', 'OS', 'NCM', 'xProd'],
          top: 5000
        });

        enrichedFiles = enrichedFiles.map(file => {
          const originalName = file.name;
          const validatedName = file.isValidated ? originalName : `Validado_${originalName}`;
          const unvalidatedName = file.isValidated ? originalName.replace(/^Validado_/i, '') : originalName;
          
          const record = history.find(h => 
            h.Title === originalName || h.Title === validatedName || h.Title === unvalidatedName
          );
          
          if (record) {
            return {
              ...file,
              nNF: record.nNF,
              CNPJ: record.CNPJ,
              OS: record.OS,
              NCM: record.NCM,
              xProd: record.xProd
            };
          }
          return file;
        });
      } catch (err) {
        console.warn('Could not enrich SP files with metadata:', err);
      }

      setSpFilesList(enrichedFiles);
    } catch (error) {
      console.error('Erro ao buscar estatísticas do SharePoint:', error);
    } finally {
      setIsFetchingSpStats(false);
    }
  }, []);

  const loadDataFromSharePoint = useCallback(async () => {
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
  }, [setRecipients, setMandatoryTags, setOsForbiddenPatterns]);

  const checkSpInitialization = useCallback(async () => {
    try {
      const recExists = await SharePointListsService.listExists('DHL_Recipients');
      const tagExists = await SharePointListsService.listExists('DHL_MandatoryTags');
      const patExists = await SharePointListsService.listExists('DHL_OSPatterns');
      const revalExists = await SharePointListsService.listExists('DHL_ValidationHistory');
      const histExists = await SharePointListsService.listExists('DHL_FullHistory');
      
      if (recExists && tagExists && patExists && revalExists && histExists) {
        setIsSpInitialized(true);
        loadDataFromSharePoint();
        return true;
      }
      return false;
    } catch (error) {
      console.error('Erro ao verificar inicialização do SharePoint:', error);
      return false;
    }
  }, [loadDataFromSharePoint]);

  const validateDateRange = useCallback((start: string, end: string) => {
    if (!start || !end) return { valid: false, message: 'Selecione as datas de início e fim.' };
    const startDate = new Date(start);
    const endDate = new Date(end);
    const diffTime = Math.abs(endDate.getTime() - startDate.getTime());
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    if (startDate > endDate) return { valid: false, message: 'A data de início não pode ser maior que a data de fim.' };
    if (diffDays > 30) return { valid: false, message: 'O intervalo máximo permitido é de 30 dias.' };
    
    return { valid: true, message: '' };
  }, []);

  const loadRevalidationFromSharePoint = useCallback(async () => {
    if (!SharePointListsService.isContextAvailable()) return;
    
    const validation = validateDateRange(revalidationStartDate, revalidationEndDate);
    if (!validation.valid) {
      showNotification('error', validation.message);
      return;
    }

    setIsFetchingRevalidation(true);
    try {
      const filter = `Created ge datetime'${revalidationStartDate}T00:00:00Z' and Created le datetime'${revalidationEndDate}T23:59:59Z'`;
      const items = await SharePointListsService.getItems('DHL_ValidationHistory', {
        select: ['Id', 'Title', 'ServerRelativeUrl', 'nNF', 'CNPJ', 'OS', 'NCM', 'xProd', 'Status', 'ValidationDate'],
        orderBy: 'Id desc',
        top: 2000,
        filter
      });
      setRevalidationItems(items);
      if (items.length === 0) {
        showNotification('success', 'Nenhum registro encontrado para este período.');
      }
    } catch (error) {
      console.error('Erro ao carregar revalidação:', error);
      showNotification('error', 'Erro ao carregar dados do SharePoint.');
    } finally {
      setIsFetchingRevalidation(false);
    }
  }, [revalidationStartDate, revalidationEndDate, validateDateRange, showNotification]);

  const loadFullHistoryFromSharePoint = useCallback(async () => {
    if (!SharePointListsService.isContextAvailable()) return;

    const validation = validateDateRange(fullHistoryStartDate, fullHistoryEndDate);
    if (!validation.valid) {
      showNotification('error', validation.message);
      return;
    }

    setIsFetchingFullHistory(true);
    try {
      const filter = `Created ge datetime'${fullHistoryStartDate}T00:00:00Z' and Created le datetime'${fullHistoryEndDate}T23:59:59Z'`;
      const items = await SharePointListsService.getItems('DHL_FullHistory', {
        select: ['Id', 'Title', 'ServerRelativeUrl', 'Status', 'nNF', 'CNPJ', 'OS', 'NCM', 'xProd', 'UserEmail', 'Source', 'ValidationDate'],
        orderBy: 'Id desc',
        top: 5000,
        filter
      });
      setFullHistory(items);
      if (items.length === 0) {
        showNotification('success', 'Nenhum registro encontrado para este período.');
      }
    } catch (error) {
      console.error('Erro ao carregar histórico:', error);
      showNotification('error', 'Erro ao carregar dados do SharePoint.');
    } finally {
      setIsFetchingFullHistory(false);
    }
  }, [fullHistoryStartDate, fullHistoryEndDate, validateDateRange, showNotification]);

  const downloadFromSharePoint = useCallback(async (serverRelativeUrl: string, fileName: string) => {
    try {
      const blob = await downloadFileFromSharePoint(serverRelativeUrl, fileName);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error(error);
      showNotification('error', 'Erro ao baixar arquivo do SharePoint.');
    }
  }, [showNotification]);

  const handleRevertSpFile = useCallback(async (spFile: { name: string; serverRelativeUrl: string }) => {
    try {
      await revertXmlFileValidation(spFile.serverRelativeUrl);
      showNotification('success', 'Validação revertida com sucesso!');
      fetchSpStats();
    } catch (error) {
      console.error(error);
      showNotification('error', 'Erro ao reverter validação.');
    }
  }, [showNotification, fetchSpStats]);

  const handleRevertValidation = useCallback(async (historyItem: any) => {
    if (!SharePointListsService.isContextAvailable()) return;
    setIsFetchingRevalidation(true);
    try {
      await revertXmlFileValidation(historyItem.ServerRelativeUrl);
      try {
        await SharePointListsService.deleteItem('DHL_ValidationHistory', historyItem.Id);
      } catch (delError) {
        console.warn('Erro ao deletar item do histórico, mas o arquivo foi restaurado:', delError);
      }
      showNotification('success', `Validação do arquivo ${historyItem.Title} revertida com sucesso!`);
    } catch (error) {
      console.error('Erro detalhado na reversão:', error);
      showNotification('error', 'Erro ao reverter validação. Verifique o console para detalhes.');
    } finally {
      setIsFetchingRevalidation(false);
    }
  }, [showNotification]);

  const syncAllToSharePoint = useCallback(async () => {
    try {
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
  }, [recipients, mandatoryTags, osForbiddenPatterns]);

  const initializeSharePoint = useCallback(async () => {
    if (!SharePointListsService.isContextAvailable()) return;
    setIsInitializingSp(true);
    try {
      // Ensure Recipients List
      await SharePointListsService.ensureList('DHL_Recipients', 'Lista de destinatários para alertas de divergência', [
        { title: 'Title', type: 'Text', required: true }
      ]);

      // Ensure Mandatory Tags List
      await SharePointListsService.ensureList('DHL_MandatoryTags', 'Tags obrigatórias para validação XML', [
        { title: 'Title', type: 'Text', required: true },
        { title: 'TagRef', type: 'Text', required: true }
      ]);

      // Ensure OS Forbidden Patterns List
      await SharePointListsService.ensureList('DHL_OSPatterns', 'Padrões proibidos no campo OS', [
        { title: 'Title', type: 'Text', required: true }
      ]);

      // Ensure Validation History List
      await SharePointListsService.ensureList('DHL_ValidationHistory', 'Histórico de validações para revalidação', [
        { title: 'Title', type: 'Text', required: true },
        { title: 'ServerRelativeUrl', type: 'Text', required: true },
        { title: 'nNF', type: 'Text' },
        { title: 'CNPJ', type: 'Text' },
        { title: 'OS', type: 'Text' },
        { title: 'NCM', type: 'Text' },
        { title: 'xProd', type: 'Text' }
      ]);

      // Ensure Full History List
      await SharePointListsService.ensureList('DHL_FullHistory', 'Histórico completo de todas as validações', [
        { title: 'Title', type: 'Text', required: true },
        { title: 'Status', type: 'Text', required: true },
        { title: 'ServerRelativeUrl', type: 'Text' },
        { title: 'nNF', type: 'Text' },
        { title: 'CNPJ', type: 'Text' },
        { title: 'OS', type: 'Text' },
        { title: 'NCM', type: 'Text' },
        { title: 'xProd', type: 'Text' },
        { title: 'UserEmail', type: 'Text' },
        { title: 'Source', type: 'Text' },
        { title: 'ValidationDate', type: 'DateTime' }
      ]);

      setIsSpInitialized(true);
      showNotification('success', 'Listas do SharePoint inicializadas com sucesso!');
      
      await syncAllToSharePoint();
      loadRevalidationFromSharePoint();
      loadFullHistoryFromSharePoint();
      
    } catch (error) {
      console.error('Erro ao inicializar SharePoint:', error);
      showNotification('error', 'Erro ao criar listas no SharePoint.');
    } finally {
      setIsInitializingSp(false);
    }
  }, [showNotification, syncAllToSharePoint, loadRevalidationFromSharePoint, loadFullHistoryFromSharePoint]);

  useEffect(() => {
    const available = SharePointListsService.isContextAvailable();
    setIsSpAvailable(available);
    if (available) {
      checkSpInitialization();
      fetchSpStats();
      const interval = setInterval(fetchSpStats, 5 * 60 * 1000);
      return () => clearInterval(interval);
    }
  }, [checkSpInitialization, fetchSpStats]);

  useEffect(() => {
    if (showSpManager && isSpAvailable) {
      fetchSpStats();
    }
  }, [showSpManager, isSpAvailable, fetchSpStats]);

  return {
    isSpAvailable,
    setIsSpAvailable,
    isSpInitialized,
    setIsSpInitialized,
    isInitializingSp,
    setIsInitializingSp,
    isFetchingSharePoint,
    setIsFetchingSharePoint,
    revalidationItems,
    setRevalidationItems,
    showRevalidation,
    setShowRevalidation,
    isFetchingRevalidation,
    revalidationSearch,
    setRevalidationSearch,
    revalidationPage,
    setRevalidationPage,
    revalidationStartDate,
    setRevalidationStartDate,
    revalidationEndDate,
    setRevalidationEndDate,
    fullHistory,
    setFullHistory,
    showFullHistory,
    setShowFullHistory,
    isFetchingFullHistory,
    fullHistorySearch,
    setFullHistorySearch,
    fullHistoryPage,
    setFullHistoryPage,
    fullHistoryStartDate,
    setFullHistoryStartDate,
    fullHistoryEndDate,
    setFullHistoryEndDate,
    spStats,
    spFilesList,
    setSpFilesList,
    isFetchingSpStats,
    showSpManager,
    setShowSpManager,
    spManagerSearch,
    setSpManagerSearch,
    spManagerPage,
    setSpManagerPage,
    spManagerStartDate,
    setSpManagerStartDate,
    spManagerEndDate,
    setSpManagerEndDate,
    filteredSpFiles,
    fetchSpStats,
    checkSpInitialization,
    loadRevalidationFromSharePoint,
    loadFullHistoryFromSharePoint,
    downloadFromSharePoint,
    handleRevertSpFile,
    handleRevertValidation,
    loadDataFromSharePoint,
    initializeSharePoint,
    syncAllToSharePoint
  };
}
