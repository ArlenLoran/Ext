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
    `<tr style="border-bottom: 1px solid #f0f0f0;">
      <td style="padding: 10px 0; vertical-align: top; width: 20px; color: #d40511; font-weight: bold;">•</td>
      <td style="padding: 10px 0; font-size: 14px; color: #444444; line-height: 1.5;">${escapeHtml(err)}</td>
    </tr>`
  ).join('');

  return `
  <!DOCTYPE html>
  <html lang="pt-BR">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>Divergência Detectada - ${fileNameSafe}</title>
    <style>
      @media only screen and (max-width: 620px) {
        .container { width: 100% !important; border-radius: 0 !important; }
        .content { padding: 20px !important; }
      }
    </style>
  </head>
  <body style="margin: 0; padding: 0; background-color: #f6f6f6; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; -webkit-font-smoothing: antialiased;">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f6f6f6;">
      <tr>
        <td align="center" style="padding: 30px 10px;">
          <table border="0" cellpadding="0" cellspacing="0" width="600" class="container" style="background-color: #ffffff; border-radius: 12px; overflow: hidden; border: 1px solid #e0e0e0;">
            
            <!-- DHL Top Bar -->
            <tr>
              <td height="8" style="background-color: #ffcc00; font-size: 1px; line-height: 8px;">&nbsp;</td>
            </tr>

            <!-- Header -->
            <tr>
              <td style="background-color: #d40511; padding: 35px 40px;" class="content">
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                  <tr>
                    <td>
                      <h1 style="color: #ffffff; margin: 0; font-size: 24px; font-weight: 900; text-transform: uppercase; letter-spacing: -0.5px; line-height: 1.2;">Divergência <br>Detectada</h1>
                    </td>
                    <td align="right" style="vertical-align: middle;">
                      <div style="background-color: #ffffff; color: #d40511; padding: 4px 10px; border-radius: 4px; font-size: 11px; font-weight: 900; text-transform: uppercase; letter-spacing: 1px;">XML Validator</div>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <!-- Main Content -->
            <tr>
              <td style="padding: 40px;" class="content">
                <p style="margin: 0 0 20px 0; font-size: 16px; color: #1a1a1a; font-weight: 700;">Olá Equipe,</p>
                <p style="margin: 0 0 30px 0; font-size: 15px; color: #555555; line-height: 1.6;">
                  O sistema identificou divergências que impedem o processamento automático do arquivo XML abaixo. Por favor, revise as informações.
                </p>

                <!-- File Info Card -->
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #fafafa; border-radius: 8px; border: 1px solid #f0f0f0; margin-bottom: 35px;">
                  <tr>
                    <td style="padding: 25px;">
                      <table border="0" cellpadding="0" cellspacing="0" width="100%">
                        <tr>
                          <td style="padding-bottom: 15px;">
                            <div style="font-size: 11px; color: #999999; text-transform: uppercase; font-weight: 800; letter-spacing: 1px; margin-bottom: 4px;">Arquivo XML</div>
                            <div style="font-size: 16px; color: #1a1a1a; font-weight: 700; word-break: break-all;">${fileNameSafe}</div>
                          </td>
                        </tr>
                        <tr>
                          <td>
                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                              <tr>
                                <td width="50%" style="border-right: 1px solid #e0e0e0; padding-right: 15px;">
                                  <div style="font-size: 11px; color: #999999; text-transform: uppercase; font-weight: 800; letter-spacing: 1px; margin-bottom: 4px;">Número (nNF)</div>
                                  <div style="font-size: 15px; color: #d40511; font-weight: 700;">${nNFSafe}</div>
                                </td>
                                <td width="50%" style="padding-left: 15px;">
                                  <div style="font-size: 11px; color: #999999; text-transform: uppercase; font-weight: 800; letter-spacing: 1px; margin-bottom: 4px;">CNPJ</div>
                                  <div style="font-size: 15px; color: #1a1a1a; font-weight: 700;">${cnpjSafe}</div>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>

                <h3 style="margin: 0 0 15px 0; font-size: 13px; color: #d40511; text-transform: uppercase; font-weight: 800; letter-spacing: 1px; border-bottom: 2px solid #ffcc00; display: inline-block; padding-bottom: 2px;">Divergências Encontradas</h3>
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 35px;">
                  ${errorsHtml}
                </table>

                <!-- CTA Button -->
                ${appUrlSafe ? `
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                  <tr>
                    <td align="center" style="padding-bottom: 35px;">
                      <a href="${appUrlSafe}" target="_blank" style="background-color: #d40511; color: #ffffff; padding: 16px 32px; text-decoration: none; font-size: 14px; font-weight: 900; border-radius: 6px; display: inline-block; text-transform: uppercase; letter-spacing: 1px; box-shadow: 0 4px 6px rgba(212, 5, 17, 0.2);">Acessar Validador</a>
                    </td>
                  </tr>
                </table>
                ` : ''}

                <!-- Notice Box -->
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #fff9f9; border-radius: 8px; border-left: 4px solid #d40511;">
                  <tr>
                    <td style="padding: 20px;">
                      <p style="margin: 0; font-size: 13px; color: #666666; line-height: 1.5;">
                        <strong style="color: #d40511;">Ação Necessária:</strong> Este arquivo não pôde ser processado automaticamente devido aos erros listados acima. Por favor, verifique e reenvie após a correção.
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <!-- Footer -->
            <tr>
              <td style="padding: 0 40px 40px 40px;" class="content">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-top: 1px solid #eeeeee; padding-top: 30px;">
                  <tr>
                    <td align="center">
                      <p style="margin: 0; font-size: 11px; color: #999999; text-transform: uppercase; font-weight: 700; letter-spacing: 1px;">DHL Logistics • XML Validation System</p>
                      <p style="margin: 12px 0 0 0; font-size: 11px; color: #cccccc;">© ${year} DHL. Todos os direitos reservados.</p>
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
      `<tr style="border-bottom: 1px dotted #f0f0f0;">
        <td style="padding: 6px 0; vertical-align: top; width: 15px; color: #d40511; font-size: 12px;">•</td>
        <td style="padding: 6px 0; font-size: 13px; color: #555555; line-height: 1.4;">${escapeHtml(err)}</td>
      </tr>`
    ).join('');

    return `
    <div style="margin-bottom: 25px; background-color: #ffffff; border-radius: 8px; border: 1px solid #e0e0e0; overflow: hidden;">
      <div style="background-color: #fcfcfc; padding: 15px 20px; border-bottom: 1px solid #f0f0f0;">
        <div style="font-size: 10px; color: #999999; text-transform: uppercase; font-weight: 800; letter-spacing: 1px; margin-bottom: 2px;">Arquivo</div>
        <div style="font-size: 14px; color: #1a1a1a; font-weight: 700; word-break: break-all;">${fileNameSafe}</div>
      </div>
      <div style="padding: 15px 20px;">
        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 15px;">
          <tr>
            <td width="50%">
              <div style="font-size: 10px; color: #999999; text-transform: uppercase; font-weight: 800; letter-spacing: 1px; margin-bottom: 2px;">nNF</div>
              <div style="font-size: 13px; color: #d40511; font-weight: 700;">${nNFSafe}</div>
            </td>
            <td width="50%">
              <div style="font-size: 10px; color: #999999; text-transform: uppercase; font-weight: 800; letter-spacing: 1px; margin-bottom: 2px;">CNPJ</div>
              <div style="font-size: 13px; color: #1a1a1a; font-weight: 700;">${cnpjSafe}</div>
            </td>
          </tr>
        </table>
        <div style="font-size: 10px; color: #d40511; text-transform: uppercase; font-weight: 800; letter-spacing: 1px; margin-bottom: 8px;">Divergências</div>
        <table border="0" cellpadding="0" cellspacing="0" width="100%">
          ${errorsHtml}
        </table>
      </div>
    </div>
    `;
  }).join('');

  return `
  <!DOCTYPE html>
  <html lang="pt-BR">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>Relatório de Divergências em Lote - ${appName}</title>
    <style>
      @media only screen and (max-width: 620px) {
        .container { width: 100% !important; border-radius: 0 !important; }
        .content { padding: 20px !important; }
      }
    </style>
  </head>
  <body style="margin: 0; padding: 0; background-color: #f6f6f6; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; -webkit-font-smoothing: antialiased;">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f6f6f6;">
      <tr>
        <td align="center" style="padding: 30px 10px;">
          <table border="0" cellpadding="0" cellspacing="0" width="600" class="container" style="background-color: #ffffff; border-radius: 12px; overflow: hidden; border: 1px solid #e0e0e0;">
            
            <!-- DHL Top Bar -->
            <tr>
              <td height="8" style="background-color: #ffcc00; font-size: 1px; line-height: 8px;">&nbsp;</td>
            </tr>

            <!-- Header -->
            <tr>
              <td style="background-color: #d40511; padding: 35px 40px;" class="content">
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                  <tr>
                    <td>
                      <h1 style="color: #ffffff; margin: 0; font-size: 24px; font-weight: 900; text-transform: uppercase; letter-spacing: -0.5px; line-height: 1.2;">Relatório <br>em Lote</h1>
                    </td>
                    <td align="right" style="vertical-align: middle;">
                      <div style="background-color: #ffffff; color: #d40511; padding: 4px 10px; border-radius: 4px; font-size: 11px; font-weight: 900; text-transform: uppercase; letter-spacing: 1px;">XML Validator</div>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <!-- Main Content -->
            <tr>
              <td style="padding: 40px;" class="content">
                <p style="margin: 0 0 20px 0; font-size: 16px; color: #1a1a1a; font-weight: 700;">Olá Equipe,</p>
                <p style="margin: 0 0 30px 0; font-size: 15px; color: #555555; line-height: 1.6;">
                  Foram identificadas divergências em <strong>${params.results.length}</strong> arquivos durante o processamento em lote. Veja os detalhes abaixo:
                </p>

                <!-- Results List -->
                ${resultsHtml}

                <!-- CTA Button -->
                ${appUrlSafe ? `
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                  <tr>
                    <td align="center" style="padding: 10px 0 35px 0;">
                      <a href="${appUrlSafe}" target="_blank" style="background-color: #d40511; color: #ffffff; padding: 16px 32px; text-decoration: none; font-size: 14px; font-weight: 900; border-radius: 6px; display: inline-block; text-transform: uppercase; letter-spacing: 1px; box-shadow: 0 4px 6px rgba(212, 5, 17, 0.2);">Acessar Validador</a>
                    </td>
                  </tr>
                </table>
                ` : ''}

                <!-- Notice Box -->
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #fcfcfc; border-radius: 8px; border-left: 4px solid #ffcc00;">
                  <tr>
                    <td style="padding: 20px;">
                      <p style="margin: 0; font-size: 13px; color: #666666; line-height: 1.5;">
                        <strong>Nota:</strong> Este é um relatório consolidado. Cada arquivo deve ser revisado individualmente no sistema.
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <!-- Footer -->
            <tr>
              <td style="padding: 0 40px 40px 40px;" class="content">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-top: 1px solid #eeeeee; padding-top: 30px;">
                  <tr>
                    <td align="center">
                      <p style="margin: 0; font-size: 11px; color: #999999; text-transform: uppercase; font-weight: 700; letter-spacing: 1px;">DHL Logistics • XML Validation System</p>
                      <p style="margin: 12px 0 0 0; font-size: 11px; color: #cccccc;">© ${year} DHL. Todos os direitos reservados.</p>
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
