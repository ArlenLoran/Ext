export interface ValidationResult {
  fileName: string;
  nNF: string;
  cnpj: string;
  ncm: string;
  osField: string;
  xProd: string;
  isValid: boolean;
  errors: string[];
  rawContent: string;
  extractedFields: Record<string, string>;
  allFields: { key: string; value: string }[];
  originalFile: File;
  sent: boolean;
  ntvStatus?: 'loading' | 'registered' | 'not_registered' | 'error';
  osStatus?: 'loading' | 'received' | 'not_received' | 'error' | 'not_found';
  ncmStatus?: 'loading' | 'registered' | 'not_registered' | 'error';
  sharepointUrl?: string;
  spValidated?: boolean;
}

export interface MandatoryTag {
  name: string;
  tag: string;
}

export interface SpFile {
  name: string;
  serverRelativeUrl: string;
  isValidated: boolean;
  timeCreated: string;
  nNF?: string;
  CNPJ?: string;
  OS?: string;
  NCM?: string;
  xProd?: string;
}

export interface SpStats {
  analyzed: number;
  pending: number;
}
