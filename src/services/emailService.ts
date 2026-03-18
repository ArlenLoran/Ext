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
    `<tr>
      <td style="padding: 10px 0; vertical-align: top; width: 20px; color: #d40511; font-weight: bold; font-family: Arial, sans-serif;">•</td>
      <td style="padding: 10px 0; font-size: 14px; color: #444444; line-height: 1.5; font-family: Arial, sans-serif; border-bottom: 1px solid #f0f0f0;">${escapeHtml(err)}</td>
    </tr>`
  ).join('');

  return `
  <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
  <html xmlns="http://www.w3.org/1999/xhtml" lang="pt-BR">
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>Divergência Detectada - ${fileNameSafe}</title>
    <!--[if mso]>
    <style type="text/css">
      body, table, td, a { font-family: Arial, Helvetica, sans-serif !important; }
    </style>
    <![endif]-->
    <style type="text/css">
      @media only screen and (max-width: 620px) {
        .container { width: 100% !important; }
        .content { padding: 20px !important; }
      }
    </style>
  </head>
  <body style="margin: 0; padding: 0; background-color: #f6f6f6; font-family: Arial, Helvetica, sans-serif;">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f6f6f6; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
      <tr>
        <td align="center" style="padding: 30px 10px;">
          <table border="0" cellpadding="0" cellspacing="0" width="600" class="container" style="background-color: #ffffff; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; border: 1px solid #e0e0e0;">
            
            <!-- DHL Top Bar -->
            <tr>
              <td height="8" style="background-color: #ffcc00; font-size: 1px; line-height: 8px;">&nbsp;</td>
            </tr>

            <!-- Header -->
            <tr>
              <td style="background-color: #d40511; padding: 35px 40px;" class="content">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                  <tr>
                    <td>
                      <h1 style="color: #ffffff; margin: 0; font-size: 24px; font-weight: bold; text-transform: uppercase; line-height: 1.2; font-family: Arial, sans-serif;">Divergência <br>Detectada</h1>
                    </td>
                    <td align="right" style="vertical-align: middle;">
                      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                        <tr>
                          <td style="background-color: #ffffff; color: #d40511; padding: 4px 10px; border-radius: 4px; font-size: 11px; font-weight: bold; text-transform: uppercase; font-family: Arial, sans-serif;">XML Validator</td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <!-- Main Content -->
            <tr>
              <td style="padding: 40px;" class="content">
                <p style="margin: 0 0 20px 0; font-size: 16px; color: #1a1a1a; font-weight: bold; font-family: Arial, sans-serif;">Olá Equipe,</p>
                <p style="margin: 0 0 30px 0; font-size: 15px; color: #555555; line-height: 1.6; font-family: Arial, sans-serif;">
                  O sistema identificou divergências que impedem o processamento automático do arquivo XML abaixo. Por favor, revise as informações.
                </p>

                <!-- File Info Card -->
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #fafafa; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; border: 1px solid #f0f0f0; margin-bottom: 35px;">
                  <tr>
                    <td style="padding: 25px;">
                      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                        <tr>
                          <td style="padding-bottom: 15px;">
                            <div style="font-size: 11px; color: #999999; text-transform: uppercase; font-weight: bold; font-family: Arial, sans-serif; margin-bottom: 4px;">Arquivo XML</div>
                            <div style="font-size: 16px; color: #1a1a1a; font-weight: bold; font-family: Arial, sans-serif; word-break: break-all;">${fileNameSafe}</div>
                          </td>
                        </tr>
                        <tr>
                          <td>
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                              <tr>
                                <td width="50%" style="border-right: 1px solid #e0e0e0; padding-right: 15px; vertical-align: top;">
                                  <div style="font-size: 11px; color: #999999; text-transform: uppercase; font-weight: bold; font-family: Arial, sans-serif; margin-bottom: 4px;">Número (nNF)</div>
                                  <div style="font-size: 15px; color: #d40511; font-weight: bold; font-family: Arial, sans-serif;">${nNFSafe}</div>
                                </td>
                                <td width="50%" style="padding-left: 15px; vertical-align: top;">
                                  <div style="font-size: 11px; color: #999999; text-transform: uppercase; font-weight: bold; font-family: Arial, sans-serif; margin-bottom: 4px;">CNPJ</div>
                                  <div style="font-size: 15px; color: #1a1a1a; font-weight: bold; font-family: Arial, sans-serif;">${cnpjSafe}</div>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>

                <h3 style="margin: 0 0 15px 0; font-size: 13px; color: #d40511; text-transform: uppercase; font-weight: bold; font-family: Arial, sans-serif; border-bottom: 2px solid #ffcc00; display: inline-block; padding-bottom: 2px;">Divergências Encontradas</h3>
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 35px; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                  ${errorsHtml}
                </table>

                <!-- CTA Button -->
                ${appUrlSafe ? `
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                  <tr>
                    <td align="center" style="padding-bottom: 35px;">
                      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                        <tr>
                          <td align="center" bgcolor="#d40511" style="border-radius: 6px;">
                            <a href="${appUrlSafe}" target="_blank" style="padding: 16px 32px; border: 1px solid #d40511; border-radius: 6px; font-family: Arial, sans-serif; font-size: 14px; color: #ffffff; text-decoration: none; font-weight: bold; display: inline-block;">Acessar Validador</a>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>
                ` : ''}

                <!-- Notice Box -->
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #fff9f9; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; border-left: 4px solid #d40511;">
                  <tr>
                    <td style="padding: 20px;">
                      <p style="margin: 0; font-size: 13px; color: #666666; line-height: 1.5; font-family: Arial, sans-serif;">
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
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-top: 1px solid #eeeeee; padding-top: 30px; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                  <tr>
                    <td align="center">
                      <p style="margin: 0; font-size: 11px; color: #999999; text-transform: uppercase; font-weight: bold; font-family: Arial, sans-serif;">DHL Logistics • XML Validation System</p>
                      <p style="margin: 12px 0 0 0; font-size: 11px; color: #cccccc; font-family: Arial, sans-serif;">© ${year} DHL. Todos os direitos reservados.</p>
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
      `<tr>
        <td style="padding: 6px 0; vertical-align: top; width: 15px; color: #d40511; font-size: 12px; font-family: Arial, sans-serif;">•</td>
        <td style="padding: 6px 0; font-size: 13px; color: #555555; line-height: 1.4; font-family: Arial, sans-serif; border-bottom: 1px dotted #f0f0f0;">${escapeHtml(err)}</td>
      </tr>`
    ).join('');

    return `
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 25px; background-color: #ffffff; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; border: 1px solid #e0e0e0;">
      <tr>
        <td style="background-color: #fcfcfc; padding: 15px 20px; border-bottom: 1px solid #f0f0f0;">
          <div style="font-size: 10px; color: #999999; text-transform: uppercase; font-weight: bold; font-family: Arial, sans-serif; margin-bottom: 2px;">Arquivo</div>
          <div style="font-size: 14px; color: #1a1a1a; font-weight: bold; font-family: Arial, sans-serif; word-break: break-all;">${fileNameSafe}</div>
        </td>
      </tr>
      <tr>
        <td style="padding: 15px 20px;">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 15px; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
            <tr>
              <td width="50%" style="vertical-align: top;">
                <div style="font-size: 10px; color: #999999; text-transform: uppercase; font-weight: bold; font-family: Arial, sans-serif; margin-bottom: 2px;">nNF</div>
                <div style="font-size: 13px; color: #d40511; font-weight: bold; font-family: Arial, sans-serif;">${nNFSafe}</div>
              </td>
              <td width="50%" style="vertical-align: top;">
                <div style="font-size: 10px; color: #999999; text-transform: uppercase; font-weight: bold; font-family: Arial, sans-serif; margin-bottom: 2px;">CNPJ</div>
                <div style="font-size: 13px; color: #1a1a1a; font-weight: bold; font-family: Arial, sans-serif;">${cnpjSafe}</div>
              </td>
            </tr>
          </table>
          <div style="font-size: 10px; color: #d40511; text-transform: uppercase; font-weight: bold; font-family: Arial, sans-serif; margin-bottom: 8px;">Divergências</div>
          <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
            ${errorsHtml}
          </table>
        </td>
      </tr>
    </table>
    `;
  }).join('');

  return `
  <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
  <html xmlns="http://www.w3.org/1999/xhtml" lang="pt-BR">
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>Relatório de Divergências em Lote - ${appName}</title>
    <!--[if mso]>
    <style type="text/css">
      body, table, td, a { font-family: Arial, Helvetica, sans-serif !important; }
    </style>
    <![endif]-->
    <style type="text/css">
      @media only screen and (max-width: 620px) {
        .container { width: 100% !important; }
        .content { padding: 20px !important; }
      }
    </style>
  </head>
  <body style="margin: 0; padding: 0; background-color: #f6f6f6; font-family: Arial, Helvetica, sans-serif;">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f6f6f6; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
      <tr>
        <td align="center" style="padding: 30px 10px;">
          <table border="0" cellpadding="0" cellspacing="0" width="600" class="container" style="background-color: #ffffff; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; border: 1px solid #e0e0e0;">
            
            <!-- DHL Top Bar -->
            <tr>
              <td height="8" style="background-color: #ffcc00; font-size: 1px; line-height: 8px;">&nbsp;</td>
            </tr>

            <!-- Header -->
            <tr>
              <td style="background-color: #d40511; padding: 35px 40px;" class="content">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                  <tr>
                    <td>
                      <h1 style="color: #ffffff; margin: 0; font-size: 24px; font-weight: bold; text-transform: uppercase; line-height: 1.2; font-family: Arial, sans-serif;">Relatório <br>em Lote</h1>
                    </td>
                    <td align="right" style="vertical-align: middle;">
                      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                        <tr>
                          <td style="background-color: #ffffff; color: #d40511; padding: 4px 10px; border-radius: 4px; font-size: 11px; font-weight: bold; text-transform: uppercase; font-family: Arial, sans-serif;">XML Validator</td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <!-- Main Content -->
            <tr>
              <td style="padding: 40px;" class="content">
                <p style="margin: 0 0 20px 0; font-size: 16px; color: #1a1a1a; font-weight: bold; font-family: Arial, sans-serif;">Olá Equipe,</p>
                <p style="margin: 0 0 30px 0; font-size: 15px; color: #555555; line-height: 1.6; font-family: Arial, sans-serif;">
                  Foram identificadas divergências em <strong>${params.results.length}</strong> arquivos durante o processamento em lote. Veja os detalhes abaixo:
                </p>

                <!-- Results List -->
                ${resultsHtml}

                <!-- CTA Button -->
                ${appUrlSafe ? `
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                  <tr>
                    <td align="center" style="padding: 10px 0 35px 0;">
                      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                        <tr>
                          <td align="center" bgcolor="#d40511" style="border-radius: 6px;">
                            <a href="${appUrlSafe}" target="_blank" style="padding: 16px 32px; border: 1px solid #d40511; border-radius: 6px; font-family: Arial, sans-serif; font-size: 14px; color: #ffffff; text-decoration: none; font-weight: bold; display: inline-block;">Acessar Validador</a>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>
                ` : ''}

                <!-- Notice Box -->
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #fcfcfc; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; border-left: 4px solid #ffcc00;">
                  <tr>
                    <td style="padding: 20px;">
                      <p style="margin: 0; font-size: 13px; color: #666666; line-height: 1.5; font-family: Arial, sans-serif;">
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
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-top: 1px solid #eeeeee; padding-top: 30px; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                  <tr>
                    <td align="center">
                      <p style="margin: 0; font-size: 11px; color: #999999; text-transform: uppercase; font-weight: bold; font-family: Arial, sans-serif;">DHL Logistics • XML Validation System</p>
                      <p style="margin: 12px 0 0 0; font-size: 11px; color: #cccccc; font-family: Arial, sans-serif;">© ${year} DHL. Todos os direitos reservados.</p>
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
