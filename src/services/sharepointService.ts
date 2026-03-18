/**
 * SharePoint Integration Service
 * Handles listing, downloading and renaming XML files from SharePoint folders.
 */

declare global {
  interface Window {
    _spPageContextInfo?: any;
  }
}

export interface SharePointXmlFile {
  name: string;
  serverRelativeUrl: string;
  file: File;
}

function getContext() {
  const ctx = window._spPageContextInfo;
  if (!ctx) {
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

function getRequestDigest(): string {
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

function buildRenamedXmlFileName(fileName: string): string {
  const trimmed = String(fileName || '').trim();
  if (!trimmed) throw new Error('Nome de arquivo inválido para renomeação.');
  if (/validado\.xml$/i.test(trimmed)) return trimmed;
  return trimmed.replace(/\.xml$/i, ' validado.xml');
}

function buildDecodedUrlApiSegment(serverRelativeUrl: string): string {
  return `decodedurl='${escapeODataString(serverRelativeUrl)}'`;
}

export async function downloadFileFromSharePoint(serverRelativeUrl: string, fileName: string): Promise<Blob> {
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

export async function listAllXmlFilesFromFolder(folderPath = 'SiteAssets/XMLs'): Promise<{ name: string; serverRelativeUrl: string; isValidated: boolean; timeCreated: string }[]> {
  const folderServerRelativeUrl = normalizeFolderServerRelativeUrl(folderPath);
  const endpoint = `${getSiteAbsoluteUrl()}/_api/web/GetFolderByServerRelativeUrl('${escapeODataString(folderServerRelativeUrl)}')/Files?$select=Name,ServerRelativeUrl,TimeCreated&$orderby=Name asc`;

  const response = await fetch(endpoint, {
    method: 'GET',
    headers: { Accept: 'application/json; odata=verbose' },
    credentials: 'same-origin'
  });

  if (!response.ok) {
    const message = await response.text().catch(() => '');
    throw new Error(message || 'Não foi possível consultar a pasta de XMLs no SharePoint.');
  }

  const data = await response.json();
  const files = (data?.d?.results || []) as Array<{ Name: string; ServerRelativeUrl: string; TimeCreated: string }>;
  
  return files
    .filter((item) => /\.xml$/i.test(item.Name))
    .map(item => ({
      name: item.Name,
      serverRelativeUrl: item.ServerRelativeUrl,
      isValidated: /validado\.xml$/i.test(item.Name),
      timeCreated: item.TimeCreated
    }));
}

export async function listXmlFilesFromFolder(folderPath = 'SiteAssets/XMLs'): Promise<SharePointXmlFile[]> {
  const folderServerRelativeUrl = normalizeFolderServerRelativeUrl(folderPath);
  const endpoint = `${getSiteAbsoluteUrl()}/_api/web/GetFolderByServerRelativeUrl('${escapeODataString(folderServerRelativeUrl)}')/Files?$select=Name,ServerRelativeUrl&$orderby=Name asc`;

  const response = await fetch(endpoint, {
    method: 'GET',
    headers: { Accept: 'application/json; odata=verbose' },
    credentials: 'same-origin'
  });

  if (!response.ok) {
    const message = await response.text().catch(() => '');
    throw new Error(message || 'Não foi possível consultar a pasta de XMLs no SharePoint.');
  }

  const data = await response.json();
  const files = (data?.d?.results || []) as Array<{ Name: string; ServerRelativeUrl: string }>;
  // Filter for .xml files and EXCLUDE those already marked as "validado"
  const xmlFiles = files.filter((item) => /\.xml$/i.test(item.Name) && !/validado\.xml$/i.test(item.Name));

  const downloaded = await Promise.all(
    xmlFiles.map(async (item) => {
      const blob = await downloadFileFromSharePoint(item.ServerRelativeUrl, item.Name);
      const file = new File([blob], item.Name, {
        type: 'text/xml',
        lastModified: Date.now()
      });

      return {
        name: item.Name,
        serverRelativeUrl: item.ServerRelativeUrl,
        file
      } satisfies SharePointXmlFile;
    })
  );

  return downloaded;
}

export async function renameXmlFileAsValidated(serverRelativeUrl: string): Promise<string> {
  const currentUrl = String(serverRelativeUrl || '').trim();
  if (!currentUrl) throw new Error('URL do arquivo no SharePoint não informada.');

  const segments = currentUrl.split('/');
  const currentName = segments.pop() || '';
  const renamed = buildRenamedXmlFileName(currentName);

  if (renamed === currentName) return currentUrl;

  const targetUrl = `${segments.join('/')}/${renamed}`;
  const endpoint = `${getSiteAbsoluteUrl()}/_api/web/GetFileByServerRelativePath(${buildDecodedUrlApiSegment(currentUrl)})/moveto(newurl='${escapeODataString(targetUrl)}',flags=1)`;

  const response = await fetch(endpoint, {
    method: 'POST',
    headers: {
      Accept: 'application/json; odata=verbose',
      'X-RequestDigest': getRequestDigest()
    },
    credentials: 'same-origin'
  });

  if (!response.ok) {
    const message = await response.text().catch(() => '');
    throw new Error(message || `Não foi possível renomear o arquivo ${currentName} no SharePoint.`);
  }

  return targetUrl;
}

export async function revertXmlFileValidation(serverRelativeUrl: string): Promise<string> {
  const currentUrl = String(serverRelativeUrl || '').trim();
  if (!currentUrl) throw new Error('URL do arquivo no SharePoint não informada.');

  const segments = currentUrl.split('/');
  const currentName = segments.pop() || '';
  
  // Remove " validado" from the end, before .xml
  const originalName = currentName.replace(/\svalidado\.xml$/i, '.xml');

  if (originalName === currentName) return currentUrl;

  const targetUrl = `${segments.join('/')}/${originalName}`;
  const endpoint = `${getSiteAbsoluteUrl()}/_api/web/GetFileByServerRelativePath(${buildDecodedUrlApiSegment(currentUrl)})/moveto(newurl='${escapeODataString(targetUrl)}',flags=1)`;

  const response = await fetch(endpoint, {
    method: 'POST',
    headers: {
      Accept: 'application/json; odata=verbose',
      'X-RequestDigest': getRequestDigest()
    },
    credentials: 'same-origin'
  });

  if (!response.ok) {
    const message = await response.text().catch(() => '');
    throw new Error(message || `Não foi possível reverter a renomeação do arquivo ${currentName} no SharePoint.`);
  }

  return targetUrl;
}
