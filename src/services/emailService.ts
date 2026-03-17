/**
 * Sends email through Power Automate HTTP trigger.
 */

const POWER_AUTOMATE_URL = 'https://51a805d34213e248a3506f5db8fe28.55.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/910aeb7e914b41efac3ac9a7888e6853/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=R04vBxsxPpahND9fKzcGoI3hDCWErCi3yrowWThXpy8';

export interface Attachment {
  Name: string;
  ContentBytes: string;
}

export async function sendEmail(
  emails: string | string[], 
  title: string, 
  bodyEmail: string,
  attachments?: Attachment[]
): Promise<any> {
  // Power Automate (Outlook) requires semicolons (;) as separators for multiple emails.
  // We join arrays with ; and also replace any commas in strings with ;
  const emailsStr = Array.isArray(emails) 
    ? emails.filter(Boolean).join(';') 
    : String(emails || '').replace(/,/g, ';');

  const data = {
    emails: emailsStr,
    Title: title,
    BodyEmail: bodyEmail,
    Attachments: attachments || []
  };

  const resp = await fetch(POWER_AUTOMATE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data)
  });

  if (!resp.ok) {
    const text = await resp.text().catch(() => '');
    throw new Error(`Falha ao enviar e-mail (${resp.status}). ${text}`);
  }

  try {
    return await resp.json();
  } catch {
    return await resp.text();
  }
}

export function buildXmlDivergenceEmailHtml(params: {
  fileName: string;
  nNF: string;
  cnpj: string;
  errors: string[];
  appUrl?: string;
}): string {
  const year = new Date().getFullYear();
  const appName = "DHL XML Validator";
  
  const escapeHtml = (s: string) =>
    String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');

  const fileNameSafe = escapeHtml(params.fileName);
  const nNFSafe = escapeHtml(params.nNF || 'N/A');
  const cnpjSafe = escapeHtml(params.cnpj || 'N/A');
  const appUrlSafe = escapeHtml(params.appUrl || '');

  const errorsHtml = params.errors.map(err => 
    `<li style="margin-bottom: 8px; color: #d40511;">${escapeHtml(err)}</li>`
  ).join('');

  return `
  <!DOCTYPE html>
  <html lang="pt-BR">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Divergência Detectada - ${fileNameSafe}</title>
  </head>
  <body style="margin: 0; padding: 0; background-color: #f4f4f4; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;">
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
      <tr>
        <td align="center" style="padding: 40px 0;">
          <table border="0" cellpadding="0" cellspacing="0" width="600" style="background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">

            <tr>
              <td height="6" style="background-color: #ffcc00; line-height: 6px; font-size: 6px;">&nbsp;</td>
            </tr>

            <tr>
              <td style="background-color: #d40511; padding: 30px 40px;">
                <h1 style="color: #ffffff; margin: 0; font-size: 22px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px;">Divergência Detectada</h1>
                <p style="margin: 8px 0 0 0; color: #ffe7ea; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px;">${appName}</p>
              </td>
            </tr>

            <tr>
              <td style="padding: 40px;">
                <p style="margin: 0 0 18px 0; font-size: 16px; color: #333333; font-weight: bold;">Olá Equipe de Logística,</p>
                <p style="margin: 0 0 22px 0; font-size: 15px; color: #555555; line-height: 1.6;">
                  O sistema <strong>${appName}</strong> identificou divergências críticas durante a validação do arquivo XML abaixo:
                </p>

                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; margin: 0 0 25px 0; background-color: #f9f9f9; border-radius: 8px; border: 1px solid #eeeeee;">
                  <tr>
                    <td style="padding: 20px;">
                      <p style="margin: 0 0 5px 0; font-size: 11px; color:#777777; text-transform: uppercase; letter-spacing: 1px; font-weight: 700;">Arquivo</p>
                      <p style="margin: 0 0 15px 0; font-size: 15px; color:#111111; font-weight: bold;">${fileNameSafe}</p>
                      
                      <table border="0" cellpadding="0" cellspacing="0" width="100%">
                        <tr>
                          <td width="50%">
                            <p style="margin: 0 0 5px 0; font-size: 11px; color:#777777; text-transform: uppercase; letter-spacing: 1px; font-weight: 700;">nNF</p>
                            <p style="margin: 0; font-size: 14px; color:#d40511; font-weight: bold;">${nNFSafe}</p>
                          </td>
                          <td width="50%">
                            <p style="margin: 0 0 5px 0; font-size: 11px; color:#777777; text-transform: uppercase; letter-spacing: 1px; font-weight: 700;">CNPJ</p>
                            <p style="margin: 0; font-size: 14px; color:#111111; font-weight: bold;">${cnpjSafe}</p>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>

                <h3 style="margin: 0 0 15px 0; font-size: 14px; color: #d40511; text-transform: uppercase; letter-spacing: 1px;">Erros Encontrados:</h3>
                <ul style="margin: 0 0 30px 0; padding-left: 20px; font-size: 14px; line-height: 1.6;">
                  ${errorsHtml}
                </ul>

                ${appUrlSafe ? `
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                  <tr>
                    <td align="center" style="padding: 10px 0 30px 0;">
                      <a href="${appUrlSafe}" target="_blank" style="background-color: #d40511; color: #ffffff; padding: 14px 28px; text-decoration: none; font-size: 13px; font-weight: 800; border-radius: 4px; display: inline-block; text-transform: uppercase; letter-spacing: 1px;">Ver no Validador</a>
                    </td>
                  </tr>
                </table>
                ` : ''}

                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #fff8f8; border-radius: 6px; border-left: 4px solid #d40511;">
                  <tr>
                    <td style="padding: 20px;">
                      <p style="margin: 0; font-size: 13px; color: #666666; line-height: 1.4;">
                        <strong>Ação Necessária:</strong> Por favor, verifique o arquivo original e realize as correções necessárias para prosseguir com o fluxo logístico.
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <tr>
              <td style="padding: 0 40px 40px 40px;">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-top: 1px solid #eeeeee; padding-top: 25px;">
                  <tr>
                    <td align="center">
                      <p style="margin: 0; font-size: 11px; color: #aaaaaa; text-transform: uppercase; letter-spacing: 1px;">DHL XML Validator • Sistema de Qualidade</p>
                      <p style="margin: 10px 0 0 0; font-size: 11px; color:#bdbdbd;">© ${year} DHL Logistics. Todos os direitos reservados.</p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

          </table>
        </td>
      </tr>
    </table>
  </body>
  </html>
  `;
}

export function buildBatchXmlDivergenceEmailHtml(params: {
  results: { fileName: string; nNF: string; cnpj: string; errors: string[] }[];
  appUrl?: string;
}): string {
  const year = new Date().getFullYear();
  const appName = "DHL XML Validator";
  
  const escapeHtml = (s: string) =>
    String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');

  const appUrlSafe = escapeHtml(params.appUrl || '');

  const resultsHtml = params.results.map(res => {
    const fileNameSafe = escapeHtml(res.fileName);
    const nNFSafe = escapeHtml(res.nNF || 'N/A');
    const cnpjSafe = escapeHtml(res.cnpj || 'N/A');
    const errorsHtml = res.errors.map(err => 
      `<li style="margin-bottom: 4px; color: #d40511;">${escapeHtml(err)}</li>`
    ).join('');

    return `
    <div style="margin-bottom: 30px; padding: 20px; background-color: #f9f9f9; border-radius: 8px; border: 1px solid #eeeeee;">
      <p style="margin: 0 0 5px 0; font-size: 11px; color:#777777; text-transform: uppercase; letter-spacing: 1px; font-weight: 700;">Arquivo</p>
      <p style="margin: 0 0 15px 0; font-size: 15px; color:#111111; font-weight: bold;">${fileNameSafe}</p>
      
      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 15px;">
        <tr>
          <td width="50%">
            <p style="margin: 0 0 5px 0; font-size: 11px; color:#777777; text-transform: uppercase; letter-spacing: 1px; font-weight: 700;">nNF</p>
            <p style="margin: 0; font-size: 14px; color:#d40511; font-weight: bold;">${nNFSafe}</p>
          </td>
          <td width="50%">
            <p style="margin: 0 0 5px 0; font-size: 11px; color:#777777; text-transform: uppercase; letter-spacing: 1px; font-weight: 700;">CNPJ</p>
            <p style="margin: 0; font-size: 14px; color:#111111; font-weight: bold;">${cnpjSafe}</p>
          </td>
        </tr>
      </table>

      <p style="margin: 0 0 8px 0; font-size: 11px; color:#d40511; text-transform: uppercase; letter-spacing: 1px; font-weight: 700;">Divergências:</p>
      <ul style="margin: 0; padding-left: 20px; font-size: 13px; line-height: 1.5;">
        ${errorsHtml}
      </ul>
    </div>
    `;
  }).join('');

  return `
  <!DOCTYPE html>
  <html lang="pt-BR">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório de Divergências em Lote - ${appName}</title>
  </head>
  <body style="margin: 0; padding: 0; background-color: #f4f4f4; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;">
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
      <tr>
        <td align="center" style="padding: 40px 0;">
          <table border="0" cellpadding="0" cellspacing="0" width="650" style="background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">

            <tr>
              <td height="6" style="background-color: #ffcc00; line-height: 6px; font-size: 6px;">&nbsp;</td>
            </tr>

            <tr>
              <td style="background-color: #d40511; padding: 30px 40px;">
                <h1 style="color: #ffffff; margin: 0; font-size: 22px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px;">Relatório de Divergências</h1>
                <p style="margin: 8px 0 0 0; color: #ffe7ea; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px;">${appName} • Processamento em Lote</p>
              </td>
            </tr>

            <tr>
              <td style="padding: 40px;">
                <p style="margin: 0 0 18px 0; font-size: 16px; color: #333333; font-weight: bold;">Olá Equipe de Logística,</p>
                <p style="margin: 0 0 25px 0; font-size: 15px; color: #555555; line-height: 1.6;">
                  Foram identificadas divergências em <strong>${params.results.length}</strong> arquivos XML durante a validação em lote. Veja os detalhes abaixo:
                </p>

                ${resultsHtml}

                ${appUrlSafe ? `
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                  <tr>
                    <td align="center" style="padding: 10px 0 30px 0;">
                      <a href="${appUrlSafe}" target="_blank" style="background-color: #d40511; color: #ffffff; padding: 14px 28px; text-decoration: none; font-size: 13px; font-weight: 800; border-radius: 4px; display: inline-block; text-transform: uppercase; letter-spacing: 1px;">Acessar Validador</a>
                    </td>
                  </tr>
                </table>
                ` : ''}

                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f9f9f9; border-radius: 6px; border-left: 4px solid #ffcc00;">
                  <tr>
                    <td style="padding: 20px;">
                      <p style="margin: 0; font-size: 13px; color: #666666; line-height: 1.4;">
                        <strong>Nota:</strong> Este é um relatório consolidado. Verifique cada arquivo individualmente para correções.
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <tr>
              <td style="padding: 0 40px 40px 40px;">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-top: 1px solid #eeeeee; padding-top: 25px;">
                  <tr>
                    <td align="center">
                      <p style="margin: 0; font-size: 11px; color: #aaaaaa; text-transform: uppercase; letter-spacing: 1px;">DHL XML Validator • Sistema de Qualidade</p>
                      <p style="margin: 10px 0 0 0; font-size: 11px; color:#bdbdbd;">© ${year} DHL Logistics. Todos os direitos reservados.</p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

          </table>
        </td>
      </tr>
    </table>
  </body>
  </html>
  `;
}
