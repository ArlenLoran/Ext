/**
 * SharePoint Integration Service
 * Handles listing, downloading and deleting PDF files from SharePoint folders using the page context.
 */

declare global {
  interface Window {
    _spPageContextInfo?: any;
  }
}

function isDev() {
  return !window._spPageContextInfo;
}

function getContext() {
  const ctx = window._spPageContextInfo;
  if (!ctx) {
    if (isDev()) {
      return {
        siteAbsoluteUrl: 'https://dhl.sharepoint.com/sites/dev',
        webServerRelativeUrl: '/sites/dev',
        formDigestValue: 'MOCK_DIGEST'
      };
    }
    throw new Error('SharePoint context (_spPageContextInfo) não encontrado. Este app deve rodar dentro de uma página SharePoint.');
  }
  return ctx;
}

function getSiteAbsoluteUrl(): string {
  return String(getContext().siteAbsoluteUrl || '').replace(/\/$/, '');
}

function getWebServerRelativeUrl(): string {
  const value = String(getContext().webServerRelativeUrl || '').trim();
  return value.endsWith('/') ? value.slice(0, -1) : value;
}

async function getRequestDigest(): Promise<string> {
  try {
    const response = await fetch(`${getSiteAbsoluteUrl()}/_api/contextinfo`, {
      method: 'POST',
      headers: { Accept: 'application/json; odata=verbose' }
    });
    if (response.ok) {
      const data = await response.json();
      return data.d.GetContextWebInformation.FormDigestValue;
    }
  } catch (err) {
    console.warn('Falha ao obter FormDigest via API, tentando contexto local:', err);
  }

  const value = String(getContext().formDigestValue || '').trim();
  if (!value) throw new Error('FormDigest não encontrado no contexto do SharePoint.');
  return value;
}

function escapeODataString(value: string): string {
  return String(value ?? '').replace(/'/g, "''");
}

function normalizeFolderServerRelativeUrl(folderPath: string): string {
  const cleanFolder = String(folderPath || '').trim().replace(/^\/+/, '').replace(/\/+$/, '');
  const webRel = getWebServerRelativeUrl();

  if (!cleanFolder) {
    throw new Error('Caminho da pasta do SharePoint não informado.');
  }

  if (cleanFolder.startsWith('/')) return cleanFolder;
  if (!webRel || webRel === '/') return `/${cleanFolder}`;
  return `${webRel}/${cleanFolder}`.replace(/\/+/g, '/');
}

function buildDecodedUrlApiSegment(serverRelativeUrl: string): string {
  const decoded = decodeURIComponent(serverRelativeUrl);
  return `decodedurl='${escapeODataString(decoded)}'`;
}

const INDEX_LIST_NAME = 'DACE_Index';

export async function ensureIndexListExists(): Promise<void> {
  if (isDev()) return;
  
  const siteUrl = getSiteAbsoluteUrl();
  const digest = await getRequestDigest();
  
  // Check if list exists
  const checkResponse = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${INDEX_LIST_NAME}')`, {
    headers: { Accept: 'application/json; odata=verbose' }
  });
  
  if (checkResponse.ok) return; // List already exists
  
  // Create list
  const createResponse = await fetch(`${siteUrl}/_api/web/lists`, {
    method: 'POST',
    headers: {
      'Accept': 'application/json; odata=verbose',
      'Content-Type': 'application/json; odata=verbose',
      'X-RequestDigest': digest
    },
    body: JSON.stringify({
      '__metadata': { 'type': 'SP.List' },
      'AllowContentTypes': true,
      'BaseTemplate': 100,
      'ContentTypesEnabled': true,
      'Description': 'Index de Notas Fiscais DACE',
      'Title': INDEX_LIST_NAME
    })
  });
  
  if (!createResponse.ok) {
    throw new Error('Falha ao criar a lista de indexação no SharePoint.');
  }
  
  // Add InvoiceNumber field
  await fetch(`${siteUrl}/_api/web/lists/getbytitle('${INDEX_LIST_NAME}')/fields`, {
    method: 'POST',
    headers: {
      'Accept': 'application/json; odata=verbose',
      'Content-Type': 'application/json; odata=verbose',
      'X-RequestDigest': digest
    },
    body: JSON.stringify({
      '__metadata': { 'type': 'SP.Field' },
      'Title': 'InvoiceNumber',
      'FieldTypeKind': 2, // Text
      'Required': false
    })
  });
}

export async function getIndexData(): Promise<Record<string, string>> {
  if (isDev()) {
    return {
      '/sites/dev/Shared Documents/DACE/DACE_NF_12345.pdf': '12345',
      '/sites/dev/Shared Documents/DACE/DACE_NF_67890.pdf': '67890'
    };
  }
  
  const siteUrl = getSiteAbsoluteUrl();
  const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${INDEX_LIST_NAME}')/items?$select=Title,InvoiceNumber&$top=5000`;
  
  const response = await fetch(endpoint, {
    headers: { Accept: 'application/json; odata=verbose' }
  });
  
  if (!response.ok) return {};
  
  const data = await response.json();
  const items = data?.d?.results || [];
  
  const index: Record<string, string> = {};
  items.forEach((item: any) => {
    index[item.Title] = item.InvoiceNumber;
  });
  
  return index;
}

export async function updateIndexItem(serverRelativeUrl: string, invoiceNumber: string): Promise<void> {
  if (isDev()) {
    console.log('Dev Mode: Updating index', serverRelativeUrl, invoiceNumber);
    return;
  }
  
  const siteUrl = getSiteAbsoluteUrl();
  const digest = await getRequestDigest();
  
  // Check if item already exists to decide between POST (add) or MERGE (update)
  const checkEndpoint = `${siteUrl}/_api/web/lists/getbytitle('${INDEX_LIST_NAME}')/items?$filter=Title eq '${escapeODataString(serverRelativeUrl)}'&$select=Id`;
  const checkResponse = await fetch(checkEndpoint, {
    headers: { Accept: 'application/json; odata=verbose' }
  });
  
  const checkData = await checkResponse.json();
  const existingItem = checkData?.d?.results?.[0];
  
  if (existingItem) {
    // Update
    await fetch(`${siteUrl}/_api/web/lists/getbytitle('${INDEX_LIST_NAME}')/items(${existingItem.Id})`, {
      method: 'POST',
      headers: {
        'Accept': 'application/json; odata=verbose',
        'Content-Type': 'application/json; odata=verbose',
        'X-RequestDigest': digest,
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      },
      body: JSON.stringify({
        '__metadata': { 'type': `SP.Data.${INDEX_LIST_NAME}ListItem` },
        'InvoiceNumber': invoiceNumber
      })
    });
  } else {
    // Add
    await fetch(`${siteUrl}/_api/web/lists/getbytitle('${INDEX_LIST_NAME}')/items`, {
      method: 'POST',
      headers: {
        'Accept': 'application/json; odata=verbose',
        'Content-Type': 'application/json; odata=verbose',
        'X-RequestDigest': digest
      },
      body: JSON.stringify({
        '__metadata': { 'type': `SP.Data.${INDEX_LIST_NAME}ListItem` },
        'Title': serverRelativeUrl,
        'InvoiceNumber': invoiceNumber
      })
    });
  }
}

export async function downloadFileFromSharePoint(serverRelativeUrl: string, fileName: string): Promise<Blob> {
  if (isDev()) {
    // Return a dummy PDF blob in dev mode
    return new Blob(['%PDF-1.4 mock content'], { type: 'application/pdf' });
  }
  const decodedUrl = buildDecodedUrlApiSegment(serverRelativeUrl);
  const endpoint = `${getSiteAbsoluteUrl()}/_api/web/GetFileByServerRelativePath(${decodedUrl})/$value`;

  const response = await fetch(endpoint, {
    method: 'GET',
    headers: {
      Accept: 'application/octet-stream'
    },
    credentials: 'same-origin'
  });

  if (!response.ok) {
    const message = await response.text().catch(() => '');
    throw new Error(message || `Não foi possível baixar o arquivo ${fileName} do SharePoint.`);
  }

  return response.blob();
}

export async function deleteFileFromSharePoint(serverRelativeUrl: string): Promise<void> {
  if (isDev()) {
    console.log('Dev Mode: Deleting file', serverRelativeUrl);
    return Promise.resolve();
  }
  const decodedUrl = buildDecodedUrlApiSegment(serverRelativeUrl);
  const endpoint = `${getSiteAbsoluteUrl()}/_api/web/GetFileByServerRelativePath(${decodedUrl})`;

  const response = await fetch(endpoint, {
    method: 'POST',
    headers: {
      'X-HTTP-Method': 'DELETE',
      'IF-MATCH': '*',
      'X-RequestDigest': await getRequestDigest()
    },
    credentials: 'same-origin'
  });

  if (!response.ok) {
    const message = await response.text().catch(() => '');
    throw new Error(message || `Não foi possível excluir o arquivo do SharePoint.`);
  }
}

export async function listPdfFilesFromFolder(folderPath = 'Shared Documents/DACE'): Promise<{ name: string; serverRelativeUrl: string; timeCreated: string; size: number }[]> {
  if (isDev()) {
    return [
      { name: 'DACE_NF_12345.pdf', serverRelativeUrl: '/sites/dev/Shared Documents/DACE/DACE_NF_12345.pdf', timeCreated: new Date().toISOString(), size: 1024 * 450 },
      { name: 'DACE_NF_67890.pdf', serverRelativeUrl: '/sites/dev/Shared Documents/DACE/DACE_NF_67890.pdf', timeCreated: new Date(Date.now() - 86400000).toISOString(), size: 1024 * 1200 },
      { name: 'DACE_NF_ABCDE.pdf', serverRelativeUrl: '/sites/dev/Shared Documents/DACE/DACE_NF_ABCDE.pdf', timeCreated: new Date(Date.now() - 172800000).toISOString(), size: 1024 * 850 },
    ];
  }
  const folderServerRelativeUrl = normalizeFolderServerRelativeUrl(folderPath);
  const endpoint = `${getSiteAbsoluteUrl()}/_api/web/GetFolderByServerRelativeUrl('${escapeODataString(folderServerRelativeUrl)}')/Files?$select=Name,ServerRelativeUrl,TimeCreated,Length&$orderby=TimeCreated desc`;

  const response = await fetch(endpoint, {
    method: 'GET',
    headers: { Accept: 'application/json; odata=verbose' },
    credentials: 'same-origin'
  });

  if (!response.ok) {
    const message = await response.text().catch(() => '');
    throw new Error(message || 'Não foi possível consultar a pasta de PDFs no SharePoint.');
  }

  const data = await response.json();
  const files = (data?.d?.results || []) as Array<{ Name: string; ServerRelativeUrl: string; TimeCreated: string; Length: string }>;
  
  return files
    .filter((item) => /\.pdf$/i.test(item.Name))
    .map(item => ({
      name: item.Name,
      serverRelativeUrl: item.ServerRelativeUrl,
      timeCreated: item.TimeCreated,
      size: parseInt(item.Length, 10)
    }));
}
