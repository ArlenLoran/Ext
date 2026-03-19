import React, { useState, useCallback } from 'react';
import { ValidationResult, MandatoryTag } from '../types';

export function useXMLValidator() {
  const [mandatoryTags, setMandatoryTags] = useState<MandatoryTag[]>(() => {
    const saved = localStorage.getItem('dhl_mandatory_tags');
    if (saved) {
      const parsed = JSON.parse(saved);
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

  const [osForbiddenPatterns, setOsForbiddenPatterns] = useState<string[]>(() => {
    const saved = localStorage.getItem('dhl_os_forbidden_patterns');
    return saved ? JSON.parse(saved) : ["OS:\\s+\\d+", "OS:\\d+[\\.,]\\d+"];
  });

  const extractXmlMetadata = useCallback((xmlText: string) => {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlText, "text/xml");
    
    const getTagValue = (tagName: string, parentPath?: string[]) => {
      let current: Element | null = xmlDoc.documentElement;
      
      if (parentPath && parentPath.length > 0) {
        for (const pTag of parentPath) {
          if (!current) break;
          const found = Array.from(current.getElementsByTagName("*")).find(el => 
            el.tagName.toLowerCase() === pTag.toLowerCase() && 
            (el.parentNode === current || el.parentNode?.parentNode === current)
          );
          if (found) {
            current = found;
          } else {
            current = null;
          }
        }
      }

      if (current) {
        const target = Array.from(current.getElementsByTagName("*")).find(el => 
          el.tagName.toLowerCase() === tagName.toLowerCase() &&
          (el.parentNode === current || el.parentNode?.parentNode === current)
        );
        if (target) {
          return target.textContent?.trim() || "";
        }
      }

      if (parentPath && parentPath.length > 0) return "";

      const allElements = xmlDoc.getElementsByTagName("*");
      for (let i = 0; i < allElements.length; i++) {
        if (allElements[i].tagName.toLowerCase() === tagName.toLowerCase()) {
          return allElements[i].textContent?.trim() || "";
        }
      }
      return "";
    };

    const nNF = getTagValue("nNF", ["infNFe", "ide"]);
    const cnpj = getTagValue("CNPJ", ["infNFe", "emit"]);
    
    const ncmElements = xmlDoc.getElementsByTagName("NCM");
    const ncmList: string[] = [];
    for (let i = 0; i < ncmElements.length; i++) {
      if (ncmElements[i].textContent) {
        ncmList.push(ncmElements[i].textContent!.trim());
      }
    }
    const ncm = ncmList.join(" | ");

    const infCpl = getTagValue("infCpl", ["infNFe", "infAdic"]);
    
    const xProdElements = xmlDoc.getElementsByTagName("xProd");
    const xProdList: string[] = [];
    for (let i = 0; i < xProdElements.length; i++) {
      if (xProdElements[i].textContent) {
        xProdList.push(xProdElements[i].textContent!.trim());
      }
    }
    const xProd = xProdList.join(" | ");

    return { nNF, cnpj, ncm, infCpl, xProd };
  }, []);

  const validateXML = useCallback(async (file: File): Promise<ValidationResult> => {
    const text = await file.text();
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(text, "text/xml");
    const metadata = extractXmlMetadata(text);
    const { nNF, cnpj, ncm, infCpl, xProd } = metadata;
    
    const errors: string[] = [];
    
    const getTagValue = (tagName: string, parentPath?: string[]) => {
      let current: Element | null = xmlDoc.documentElement;
      
      if (parentPath && parentPath.length > 0) {
        for (const pTag of parentPath) {
          if (!current) break;
          const found = Array.from(current.getElementsByTagName("*")).find(el => 
            el.tagName.toLowerCase() === pTag.toLowerCase() && 
            (el.parentNode === current || el.parentNode?.parentNode === current)
          );
          if (found) {
            current = found;
          } else {
            current = null;
          }
        }
      }

      if (current) {
        const target = Array.from(current.getElementsByTagName("*")).find(el => 
          el.tagName.toLowerCase() === tagName.toLowerCase() &&
          (el.parentNode === current || el.parentNode?.parentNode === current)
        );
        if (target) {
          return target.textContent?.trim() || "";
        }
      }

      if (parentPath && parentPath.length > 0) return "";

      const allElements = xmlDoc.getElementsByTagName("*");
      for (let i = 0; i < allElements.length; i++) {
        if (allElements[i].tagName.toLowerCase() === tagName.toLowerCase()) {
          return allElements[i].textContent?.trim() || "";
        }
      }
      return "";
    };

    mandatoryTags.forEach(m => {
      let val = "";
      if (m.tag.toLowerCase() === 'cnpj') {
        val = getTagValue("CNPJ", ["infNFe", "emit"]);
      } else if (m.tag.toLowerCase() === 'nnf') {
        val = getTagValue("nNF", ["infNFe", "ide"]);
      } else {
        val = getTagValue(m.tag);
      }

      if (!val) {
        errors.push(`Campo obrigatório '${m.name}' não encontrado ou vazio.`);
      }
    });

    const osMatch = infCpl.match(/OS:(\d+)/);
    const osValue = osMatch ? osMatch[0] : "";
    
    if (!osValue) {
      if (infCpl.toLowerCase().includes("os:")) {
        errors.push("Campo OS encontrado mas em formato inválido (deve ser 'OS:12345678' sem espaços ou pontos).");
      } else {
        errors.push("Campo OS não encontrado nas informações complementares (infCpl).");
      }
    } else {
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

    const allFields: { key: string; value: string }[] = [];
    const extractedFields: Record<string, string> = {};
    
    extractedFields["nNF"] = getTagValue("nNF", ["infNFe", "ide"]);
    extractedFields["CNPJ"] = getTagValue("CNPJ", ["infNFe", "emit"]);
    
    const traverse = (node: Node) => {
      if (node.nodeType === 1) {
        const element = node as Element;
        if (element.children.length === 0 && element.textContent?.trim()) {
          const tag = element.tagName;
          const val = element.textContent.trim();
          
          if (!extractedFields[tag]) {
            extractedFields[tag] = val;
          }
          
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
      xProd,
      isValid: errors.length === 0,
      errors,
      rawContent: text,
      extractedFields,
      allFields,
      originalFile: file,
      sent: false
    };
  }, [mandatoryTags, osForbiddenPatterns, extractXmlMetadata]);

  const checkNtvStatus = useCallback(async (fileName: string, ncm: string, setResults: React.Dispatch<React.SetStateAction<ValidationResult[]>>) => {
    if (!ncm) return;
    
    setResults(prev => {
      const updated = [...prev];
      const idx = updated.findIndex(r => r.fileName === fileName);
      if (idx !== -1) {
        updated[idx] = { ...updated[idx], ntvStatus: 'loading' };
      }
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
        const idx = updated.findIndex(r => r.fileName === fileName);
        if (idx !== -1) {
          updated[idx] = { ...updated[idx], ntvStatus: isRegistered ? 'registered' : 'not_registered' };
        }
        return updated;
      });
    } catch (error) {
      console.error("Erro ao verificar NTV:", error);
      setResults(prev => {
        const updated = [...prev];
        const idx = updated.findIndex(r => r.fileName === fileName);
        if (idx !== -1) {
          updated[idx] = { ...updated[idx], ntvStatus: 'error' };
        }
        return updated;
      });
    }
  }, []);

  return {
    mandatoryTags,
    setMandatoryTags,
    osForbiddenPatterns,
    setOsForbiddenPatterns,
    validateXML,
    extractXmlMetadata,
    checkNtvStatus
  };
}
