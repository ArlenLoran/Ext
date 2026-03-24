import { useState, useCallback, useEffect, useMemo } from 'react';
import { listXmlFilesFromFolder, renameXmlFileAsValidated, revertXmlFileValidation, downloadFileFromSharePoint, listAllXmlFilesFromFolder } from '../services/sharepointService';
import { dbService } from '../services/dbService';
import { SpFile, SpStats, MandatoryTag } from '../types';

export function useSharePointManager(
  showNotification: (type: 'success' | 'error', message: string) => void,
  recipients: string[],
  setRecipients: (r: string[]) => void,
  mandatoryTags: MandatoryTag[],
  setMandatoryTags: (t: MandatoryTag[]) => void,
  osForbiddenPatterns: string[],
  setOsForbiddenPatterns: (p: string[]) => void,
  importLimit: number,
  setImportLimit: (l: number) => void
) {
  const [isSpAvailable, setIsSpAvailable] = useState(true);
  const [isSpInitialized, setIsSpInitialized] = useState(true);
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
    setIsFetchingSpStats(true);
    try {
      const allFiles = await listAllXmlFilesFromFolder();
      
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
        const history = await dbService.getHistory();

        enrichedFiles = enrichedFiles.map(file => {
          const originalName = file.name;
          const validatedName = file.isValidated ? originalName : `Validado_${originalName}`;
          const unvalidatedName = file.isValidated ? originalName.replace(/^Validado_/i, '') : originalName;
          
          const record = history.find((h: any) => 
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
      const spRecipients = await dbService.getRecipients();
      if (spRecipients.length > 0) {
        setRecipients(spRecipients.map((item: any) => item.Title));
      }

      const spTags = await dbService.getTags();
      if (spTags.length > 0) {
        setMandatoryTags(spTags.map((item: any) => ({ name: item.Title, tag: item.TagRef })));
      }

      const spPatterns = await dbService.getOSPatterns();
      if (spPatterns.length > 0) {
        setOsForbiddenPatterns(spPatterns.map((item: any) => item.Title));
      }

      const spConfig = await dbService.getConfig();
      const limitConfig = spConfig.find((c: any) => c.Title === 'ImportLimit');
      if (limitConfig) {
        setImportLimit(parseInt(limitConfig.Value, 10));
      }
    } catch (error) {
      console.error('Erro ao carregar dados do banco de dados:', error);
    }
  }, [setRecipients, setMandatoryTags, setOsForbiddenPatterns, setImportLimit]);

  const checkSpInitialization = useCallback(async () => {
    // With backend, we assume it's always initialized or handled by backend
    setIsSpInitialized(true);
    loadDataFromSharePoint();
    return true;
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
    const validation = validateDateRange(revalidationStartDate, revalidationEndDate);
    if (!validation.valid) {
      showNotification('error', validation.message);
      return;
    }

    setIsFetchingRevalidation(true);
    try {
      const items = await dbService.getValidationHistory();
      // Filter by date range manually if backend doesn't support it yet
      const filtered = items.filter((item: any) => {
        const date = new Date(item.ValidationDate);
        return date >= new Date(revalidationStartDate) && date <= new Date(revalidationEndDate + 'T23:59:59Z');
      });
      setRevalidationItems(filtered);
      if (filtered.length === 0) {
        showNotification('success', 'Nenhum registro encontrado para este período.');
      }
    } catch (error) {
      console.error('Erro ao carregar revalidação:', error);
      showNotification('error', 'Erro ao carregar dados do banco de dados.');
    } finally {
      setIsFetchingRevalidation(false);
    }
  }, [revalidationStartDate, revalidationEndDate, validateDateRange, showNotification]);

  const loadFullHistoryFromSharePoint = useCallback(async () => {
    const validation = validateDateRange(fullHistoryStartDate, fullHistoryEndDate);
    if (!validation.valid) {
      showNotification('error', validation.message);
      return;
    }

    setIsFetchingFullHistory(true);
    try {
      const items = await dbService.getHistory();
      const filtered = items.filter((item: any) => {
        const date = new Date(item.ValidationDate);
        return date >= new Date(fullHistoryStartDate) && date <= new Date(fullHistoryEndDate + 'T23:59:59Z');
      });
      setFullHistory(filtered);
      if (filtered.length === 0) {
        showNotification('success', 'Nenhum registro encontrado para este período.');
      }
    } catch (error) {
      console.error('Erro ao carregar histórico:', error);
      showNotification('error', 'Erro ao carregar dados do banco de dados.');
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
      await revertXmlFileValidation(spFile.serverRelativeUrl, spFile.name);
      showNotification('success', 'Validação revertida com sucesso!');
      fetchSpStats();
    } catch (error) {
      console.error(error);
      showNotification('error', 'Erro ao reverter validação.');
    }
  }, [showNotification, fetchSpStats]);

  const handleRevertValidation = useCallback(async (historyItem: any) => {
    setIsFetchingRevalidation(true);
    try {
      await revertXmlFileValidation(historyItem.ServerRelativeUrl, historyItem.Title);
      try {
        await dbService.deleteValidationHistory(historyItem.ID);
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
        await dbService.addRecipient(email);
      }
      for (const tag of mandatoryTags) {
        await dbService.addTag(tag.name, tag.tag);
      }
      for (const pattern of osForbiddenPatterns) {
        await dbService.addOSPattern(pattern);
      }
      await dbService.saveConfig('ImportLimit', importLimit.toString());
    } catch (error) {
      console.error('Erro ao sincronizar dados com banco de dados:', error);
    }
  }, [recipients, mandatoryTags, osForbiddenPatterns, importLimit]);

  const initializeSharePoint = useCallback(async () => {
    setIsInitializingSp(true);
    try {
      await dbService.initializeDb();
      setIsSpInitialized(true);
      showNotification('success', 'Integração com banco de dados inicializada!');
      
      await syncAllToSharePoint();
      loadRevalidationFromSharePoint();
      loadFullHistoryFromSharePoint();
      
    } catch (error) {
      console.error('Erro ao inicializar banco de dados:', error);
      showNotification('error', 'Erro ao inicializar integração.');
    } finally {
      setIsInitializingSp(false);
    }
  }, [showNotification, syncAllToSharePoint, loadRevalidationFromSharePoint, loadFullHistoryFromSharePoint]);

  useEffect(() => {
    checkSpInitialization();
    fetchSpStats();
    const interval = setInterval(fetchSpStats, 5 * 60 * 1000);
    return () => clearInterval(interval);
  }, [checkSpInitialization, fetchSpStats]);

  useEffect(() => {
    if (showSpManager) {
      fetchSpStats();
    }
  }, [showSpManager, fetchSpStats]);

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
