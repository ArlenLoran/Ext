/**
 * Database Integration Service
 * Handles CRUD operations for recipients, tags, OS patterns, and history via the backend API.
 */

export interface Recipient {
  Title: string;
}

export interface MandatoryTag {
  Title: string;
  TagRef: string;
}

export interface OSPattern {
  Title: string;
}

export interface Config {
  Title: string;
  Value: string;
}

export interface ValidationHistory {
  ID?: number;
  Title: string;
  ServerRelativeUrl: string;
  nNF: string;
  CNPJ: string;
  OS: string;
  NCM: string;
  xProd: string;
  Status: string;
  ValidationDate: string;
}

export interface FullHistory extends ValidationHistory {
  UserEmail: string;
  Source: string;
}

const apiFetch = async (url: string, options?: RequestInit) => {
  const response = await fetch(url, options);
  if (!response.ok) {
    const error = await response.json().catch(() => ({ error: 'Unknown error' }));
    throw new Error(error.error || `Request failed: ${response.statusText}`);
  }
  return response.json();
};

export const dbService = {
  // Recipients
  getRecipients: () => apiFetch('/api/db/recipients'),
  addRecipient: (Title: string) => apiFetch('/api/db/recipients', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ Title })
  }),
  deleteRecipient: (Title: string) => apiFetch(`/api/db/recipients/${encodeURIComponent(Title)}`, {
    method: 'DELETE'
  }),
  updateRecipient: (oldTitle: string, newTitle: string) => apiFetch(`/api/db/recipients/${encodeURIComponent(oldTitle)}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ Title: newTitle })
  }),

  // Tags
  getTags: () => apiFetch('/api/db/tags'),
  addTag: (Title: string, TagRef: string) => apiFetch('/api/db/tags', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ Title, TagRef })
  }),
  deleteTag: (TagRef: string) => apiFetch(`/api/db/tags/${encodeURIComponent(TagRef)}`, {
    method: 'DELETE'
  }),
  updateTag: (TagRef: string, Title: string) => apiFetch(`/api/db/tags/${encodeURIComponent(TagRef)}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ Title })
  }),

  // OS Patterns
  getOSPatterns: () => apiFetch('/api/db/os-patterns'),
  addOSPattern: (Title: string) => apiFetch('/api/db/os-patterns', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ Title })
  }),
  deleteOSPattern: (Title: string) => apiFetch(`/api/db/os-patterns/${encodeURIComponent(Title)}`, {
    method: 'DELETE'
  }),

  // Config
  getConfig: () => apiFetch('/api/db/config'),
  saveConfig: (Title: string, Value: string) => apiFetch('/api/db/config', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ Title, Value })
  }),

  // History
  getHistory: () => apiFetch('/api/db/history'),
  addHistory: (data: FullHistory) => apiFetch('/api/db/history', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data)
  }),

  // Validation History
  getValidationHistory: () => apiFetch('/api/db/validation-history'),
  addValidationHistory: (data: ValidationHistory) => apiFetch('/api/db/validation-history', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data)
  }),
  deleteValidationHistory: (id: number) => apiFetch(`/api/db/validation-history/${id}`, {
    method: 'DELETE'
  }),

  // Registered Products
  getRegisteredProducts: (): Promise<string[]> => apiFetch('/api/db/registered-products'),
  addRegisteredProduct: (productName: string) => apiFetch('/api/db/registered-products', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ productName })
  }),
  deleteRegisteredProduct: (productName: string) => apiFetch(`/api/db/registered-products/${encodeURIComponent(productName)}`, {
    method: 'DELETE'
  }),

  // External DB Queries
  queryNtv: (product: string) => apiFetch('/api/db/query/ntv', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ product })
  }),
  queryOs: (osNumber: string) => apiFetch('/api/db/query/os', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ osNumber })
  }),
  queryNcm: (ncm: string) => apiFetch('/api/db/query/ncm', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ncm })
  }),

  // Initialization
  initializeDb: () => apiFetch('/api/db/initialize', {
    method: 'POST'
  })
};
