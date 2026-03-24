/**
 * SharePoint Integration Service (Backend Proxy)
 * Handles listing, downloading and renaming XML files via the Next.js backend.
 */

export interface SharePointXmlFile {
  name: string;
  serverRelativeUrl: string;
  file: File;
}

export async function downloadFileFromSharePoint(itemId: string, fileName: string): Promise<Blob> {
  const response = await fetch(`/api/sharepoint/download/${itemId}`);

  if (!response.ok) {
    const message = await response.text().catch(() => '');
    throw new Error(message || `Não foi possível baixar o arquivo ${fileName} do SharePoint.`);
  }

  return response.blob();
}

export async function listAllXmlFilesFromFolder(): Promise<{ name: string; serverRelativeUrl: string; isValidated: boolean; timeCreated: string }[]> {
  const response = await fetch('/api/sharepoint/list');

  if (!response.ok) {
    const message = await response.text().catch(() => '');
    throw new Error(message || 'Não foi possível consultar a pasta de XMLs no SharePoint.');
  }

  const files = await response.json();
  
  return files.map((item: any) => ({
    name: item.name,
    serverRelativeUrl: item.id,
    isValidated: /^Validado_/i.test(item.name),
    timeCreated: item.timeCreated
  }));
}

export async function listXmlFilesFromFolder(): Promise<SharePointXmlFile[]> {
  const files = await listAllXmlFilesFromFolder();
  
  // Filter for .xml files and EXCLUDE those already marked as "Validado" (prefix Validado_)
  const xmlFiles = files.filter((item) => /\.xml$/i.test(item.name) && !item.isValidated);

  const downloaded = await Promise.all(
    xmlFiles.map(async (item) => {
      const blob = await downloadFileFromSharePoint(item.serverRelativeUrl, item.name);
      const file = new File([blob], item.name, {
        type: 'text/xml',
        lastModified: Date.now()
      });

      return {
        name: item.name,
        serverRelativeUrl: item.serverRelativeUrl,
        file
      } satisfies SharePointXmlFile;
    })
  );

  return downloaded;
}

export function buildRenamedXmlFileName(originalName: string): string {
  if (/^Validado_/i.test(originalName)) return originalName;
  return `Validado_${originalName}`;
}

export async function renameXmlFileAsValidated(itemId: string, fileName: string): Promise<string> {
  const newName = buildRenamedXmlFileName(fileName);
  const response = await fetch('/api/sharepoint/rename', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ itemId, newName })
  });

  if (!response.ok) {
    const message = await response.text().catch(() => '');
    throw new Error(message || `Não foi possível renomear o arquivo ${fileName} no SharePoint.`);
  }

  return newName;
}

export async function revertXmlFileValidation(itemId: string, fileName: string): Promise<string> {
  const newName = fileName.replace(/^Validado_/i, '');
  const response = await fetch('/api/sharepoint/rename', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ itemId, newName })
  });

  if (!response.ok) {
    const message = await response.text().catch(() => '');
    throw new Error(message || `Não foi possível reverter a validação do arquivo ${fileName} no SharePoint.`);
  }

  return newName;
}

export async function deleteXmlFile(itemId: string): Promise<void> {
  const response = await fetch(`/api/sharepoint/delete/${itemId}`, {
    method: 'DELETE'
  });

  if (!response.ok) {
    const message = await response.text().catch(() => '');
    throw new Error(message || `Não foi possível excluir o arquivo do SharePoint.`);
  }
}
