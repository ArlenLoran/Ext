import { useState, useCallback } from 'react';
import { ValidationResult } from '../types';

export function useResults() {
  const [results, setResults] = useState<ValidationResult[]>([]);
  const [expandedIndices, setExpandedIndices] = useState<number[]>([]);
  const [resultsFilters, setResultsFilters] = useState({
    nNF: '',
    cnpj: '',
    ncm: '',
    os: '',
    xProd: ''
  });

  const clearAll = useCallback(() => setResults([]), []);

  const toggleExpand = useCallback((index: number) => {
    setExpandedIndices(prev => 
      prev.includes(index) ? prev.filter(i => i !== index) : [...prev, index]
    );
  }, []);

  const removeResult = useCallback((index: number) => {
    setResults(prev => prev.filter((_, i) => i !== index));
  }, []);

  return {
    results,
    setResults,
    expandedIndices,
    setExpandedIndices,
    resultsFilters,
    setResultsFilters,
    clearAll,
    toggleExpand,
    removeResult
  };
}
