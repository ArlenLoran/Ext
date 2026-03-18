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
      <td style="padding: 12px 16px; font-size: 14px; color: #4b5563; line-height: 1.5; font-family: Arial, sans-serif; border-bottom: 1px solid #f3f4f6; background-color: #ffffff;">
        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
          <tr>
            <td style="vertical-align: top; width: 24px; padding-top: 2px;">
              <table border="0" cellpadding="0" cellspacing="0" width="16" height="16" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                <tr>
                  <td align="center" bgcolor="#d40511" style="border-radius: 4px; font-size: 10px; color: #ffffff; font-weight: bold; font-family: Arial, sans-serif;">!</td>
                </tr>
              </table>
            </td>
            <td style="padding-left: 8px; color: #4b5563; font-family: Arial, sans-serif;">${escapeHtml(err)}</td>
          </tr>
        </table>
      </td>
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
        .content { padding: 24px !important; }
      }
    </style>
  </head>
  <body style="margin: 0; padding: 0; background-color: #f8f9fa; font-family: Arial, Helvetica, sans-serif;">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8f9fa; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
      <tr>
        <td align="center" style="padding: 40px 10px;">
          <table border="0" cellpadding="0" cellspacing="0" width="600" class="container" style="background-color: #ffffff; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; border-radius: 24px; border: 1px solid #e5e7eb; overflow: hidden;">
            
            <!-- Header Section -->
            <tr>
              <td style="background-color: #1a1a1a; padding: 0;">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                  <tr>
                    <td height="6" bgcolor="#d40511" style="font-size: 1px; line-height: 6px;">&nbsp;</td>
                  </tr>
                  <tr>
                    <td style="padding: 30px 40px;" class="content">
                      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                        <tr>
                          <td>
                            <div style="color: #ffcc00; font-size: 10px; font-weight: 900; text-transform: uppercase; letter-spacing: 2px; font-family: Arial, sans-serif; margin-bottom: 8px;">ALERTA DE SISTEMA</div>
                            <h1 style="color: #ffffff; margin: 0; font-size: 28px; font-weight: bold; text-transform: uppercase; line-height: 1; font-family: Arial, sans-serif; letter-spacing: -1px;">Divergência <br><span style="color: #d40511;">Detectada</span></h1>
                          </td>
                          <td align="right" style="vertical-align: middle;">
                            <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                              <tr>
                                <td style="border: 1px solid #333333; padding: 8px 12px; border-radius: 8px; background-color: #262626;">
                                  <div style="color: #ffffff; font-size: 10px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; font-family: Arial, sans-serif;">XML Validator</div>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <!-- Body Content -->
            <tr>
              <td style="padding: 40px;" class="content">
                <p style="margin: 0 0 16px 0; font-size: 18px; color: #111827; font-weight: bold; font-family: Arial, sans-serif;">Olá Equipe,</p>
                <p style="margin: 0 0 32px 0; font-size: 15px; color: #4b5563; line-height: 1.6; font-family: Arial, sans-serif;">
                  O sistema identificou divergências críticas que impedem o processamento automático do arquivo XML. Por favor, revise os detalhes abaixo.
                </p>

                <!-- File Info Card -->
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #ffffff; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; border: 1px solid #f3f4f6; border-radius: 16px; margin-bottom: 48px;">
                  <tr>
                    <td style="padding: 24px; border-radius: 16px; border: 1px solid #f3f4f6;">
                      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                        <tr>
                          <td style="padding-bottom: 20px;">
                            <div style="font-size: 10px; color: #9ca3af; text-transform: uppercase; font-weight: 900; letter-spacing: 1px; font-family: Arial, sans-serif; margin-bottom: 6px;">IDENTIFICAÇÃO DO ARQUIVO</div>
                            <div style="font-size: 16px; color: #111827; font-weight: bold; font-family: Arial, sans-serif; word-break: break-all;">${fileNameSafe}</div>
                          </td>
                        </tr>
                        <tr>
                          <td>
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                              <tr>
                                <td width="50%" style="vertical-align: top;">
                                  <div style="font-size: 10px; color: #9ca3af; text-transform: uppercase; font-weight: 900; letter-spacing: 1px; font-family: Arial, sans-serif; margin-bottom: 6px;">NÚMERO (nNF)</div>
                                  <div style="font-size: 15px; color: #d40511; font-weight: bold; font-family: Arial, sans-serif;">${nNFSafe}</div>
                                </td>
                                <td width="50%" style="vertical-align: top; padding-left: 20px; border-left: 1px solid #f3f4f6;">
                                  <div style="font-size: 10px; color: #9ca3af; text-transform: uppercase; font-weight: 900; letter-spacing: 1px; font-family: Arial, sans-serif; margin-bottom: 6px;">CNPJ EMISSOR</div>
                                  <div style="font-size: 15px; color: #111827; font-weight: bold; font-family: Arial, sans-serif;">${cnpjSafe}</div>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>

                <!-- Errors Table Section -->
                <div style="margin-bottom: 12px;">
                  <span style="font-size: 12px; color: #111827; text-transform: uppercase; font-weight: 900; letter-spacing: 1px; font-family: Arial, sans-serif;">Divergências Encontradas</span>
                </div>
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; border: 1px solid #f3f4f6; border-radius: 12px; overflow: hidden; margin-bottom: 48px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                  <thead>
                    <tr>
                      <th bgcolor="#f9fafb" align="left" style="padding: 12px 16px; border-bottom: 2px solid #f3f4f6; font-size: 10px; color: #6b7280; text-transform: uppercase; font-weight: bold; letter-spacing: 1px; font-family: Arial, sans-serif;">Descrição do Erro</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${errorsHtml}
                  </tbody>
                </table>

                <!-- Notice Box -->
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f9fafb; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; border-radius: 12px; border: 1px solid #f3f4f6;">
                  <tr>
                    <td style="padding: 24px;">
                      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                        <tr>
                          <td style="vertical-align: top; width: 20px; padding-right: 12px;">
                            <div style="color: #d40511; font-size: 18px; font-weight: bold;">ℹ</div>
                          </td>
                          <td>
                            <p style="margin: 0; font-size: 13px; color: #6b7280; line-height: 1.6; font-family: Arial, sans-serif;">
                              <strong style="color: #111827;">Ação Necessária:</strong> Este arquivo não pôde ser processado automaticamente. Por favor, verifique as divergências listadas e reenvie o arquivo após a correção.
                            </p>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <!-- Footer -->
            <tr>
              <td style="padding: 0 40px 40px 40px;" class="content">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-top: 1px solid #f3f4f6; padding-top: 32px; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                  <tr>
                    <td align="center">
                      <div style="margin-bottom: 16px;">
                        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                          <tr>
                            <td bgcolor="#d40511" width="30" height="4" style="font-size: 1px; line-height: 4px;">&nbsp;</td>
                            <td bgcolor="#ffcc00" width="30" height="4" style="font-size: 1px; line-height: 4px;">&nbsp;</td>
                          </tr>
                        </table>
                      </div>
                      <p style="margin: 0; font-size: 11px; color: #9ca3af; text-transform: uppercase; font-weight: 900; letter-spacing: 2px; font-family: Arial, sans-serif;">DHL Logistics • XML Validation System</p>
                      <p style="margin: 12px 0 0 0; font-size: 11px; color: #d1d5db; font-family: Arial, sans-serif;">© ${year} DHL. Todos os direitos reservados.</p>
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
        <td style="padding: 12px 0; vertical-align: top; width: 24px;">
          <table border="0" cellpadding="0" cellspacing="0" width="16" height="16" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
            <tr>
              <td align="center" bgcolor="#d40511" style="border-radius: 4px; font-size: 10px; color: #ffffff; font-weight: bold; font-family: Arial, sans-serif;">!</td>
            </tr>
          </table>
        </td>
        <td style="padding: 12px 0; font-size: 14px; color: #4b5563; line-height: 1.5; font-family: Arial, sans-serif; border-bottom: 1px solid #f3f4f6;">${escapeHtml(err)}</td>
      </tr>`
    ).join('');

    return `
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 32px; background-color: #ffffff; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; border: 1px solid #f3f4f6; border-radius: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
      <tr>
        <td style="background-color: #f9fafb; padding: 20px 24px; border-bottom: 1px solid #f3f4f6; border-radius: 16px 16px 0 0;">
          <div style="font-size: 10px; color: #9ca3af; text-transform: uppercase; font-weight: 900; letter-spacing: 1px; font-family: Arial, sans-serif; margin-bottom: 6px;">ARQUIVO</div>
          <div style="font-size: 15px; color: #111827; font-weight: bold; font-family: Arial, sans-serif; word-break: break-all;">${fileNameSafe}</div>
        </td>
      </tr>
      <tr>
        <td style="padding: 24px;">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-bottom: 24px; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
            <tr>
              <td width="50%" style="vertical-align: top;">
                <div style="font-size: 10px; color: #9ca3af; text-transform: uppercase; font-weight: 900; letter-spacing: 1px; font-family: Arial, sans-serif; margin-bottom: 6px;">NÚMERO (nNF)</div>
                <div style="font-size: 14px; color: #d40511; font-weight: bold; font-family: Arial, sans-serif;">${nNFSafe}</div>
              </td>
              <td width="50%" style="vertical-align: top; padding-left: 20px; border-left: 1px solid #f3f4f6;">
                <div style="font-size: 10px; color: #9ca3af; text-transform: uppercase; font-weight: 900; letter-spacing: 1px; font-family: Arial, sans-serif; margin-bottom: 6px;">CNPJ EMISSOR</div>
                <div style="font-size: 14px; color: #111827; font-weight: bold; font-family: Arial, sans-serif;">${cnpjSafe}</div>
              </td>
            </tr>
          </table>
          
          <div style="margin-bottom: 8px;">
            <span style="font-size: 11px; color: #111827; text-transform: uppercase; font-weight: 900; letter-spacing: 1px; font-family: Arial, sans-serif;">Divergências</span>
          </div>
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
        .content { padding: 24px !important; }
      }
    </style>
  </head>
  <body style="margin: 0; padding: 0; background-color: #f8f9fa; font-family: Arial, Helvetica, sans-serif;">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f8f9fa; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
      <tr>
        <td align="center" style="padding: 40px 10px;">
          <table border="0" cellpadding="0" cellspacing="0" width="600" class="container" style="background-color: #ffffff; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; border-radius: 24px; border: 1px solid #e5e7eb; overflow: hidden;">
            
            <!-- Header Section -->
            <tr>
              <td style="background-color: #1a1a1a; padding: 0;">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                  <tr>
                    <td height="6" bgcolor="#d40511" style="font-size: 1px; line-height: 6px;">&nbsp;</td>
                  </tr>
                  <tr>
                    <td style="padding: 30px 40px;" class="content">
                      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                        <tr>
                          <td>
                            <div style="color: #ffcc00; font-size: 10px; font-weight: 900; text-transform: uppercase; letter-spacing: 2px; font-family: Arial, sans-serif; margin-bottom: 8px;">RELATÓRIO CONSOLIDADO</div>
                            <h1 style="color: #ffffff; margin: 0; font-size: 28px; font-weight: bold; text-transform: uppercase; line-height: 1; font-family: Arial, sans-serif; letter-spacing: -1px;">Processamento <br><span style="color: #d40511;">em Lote</span></h1>
                          </td>
                          <td align="right" style="vertical-align: middle;">
                            <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                              <tr>
                                <td style="border: 1px solid #333333; padding: 8px 12px; border-radius: 8px; background-color: #262626;">
                                  <div style="color: #ffffff; font-size: 10px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; font-family: Arial, sans-serif;">XML Validator</div>
                                </td>
                              </tr>
                            </table>
                          </td>
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
                <p style="margin: 0 0 16px 0; font-size: 18px; color: #111827; font-weight: bold; font-family: Arial, sans-serif;">Olá Equipe,</p>
                <p style="margin: 0 0 32px 0; font-size: 15px; color: #4b5563; line-height: 1.6; font-family: Arial, sans-serif;">
                  Foram identificadas divergências em <strong>${params.results.length}</strong> arquivos durante o processamento em lote. Veja os detalhes abaixo:
                </p>

                <!-- Results List -->
                ${resultsHtml}

                <!-- Notice Box -->
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f9fafb; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; border-radius: 12px; border: 1px solid #f3f4f6;">
                  <tr>
                    <td style="padding: 24px;">
                      <p style="margin: 0; font-size: 13px; color: #6b7280; line-height: 1.6; font-family: Arial, sans-serif;">
                        <strong style="color: #111827;">Nota:</strong> Este é um relatório consolidado. Cada arquivo deve ser revisado individualmente no sistema para garantir a conformidade.
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <!-- Footer -->
            <tr>
              <td style="padding: 0 40px 40px 40px;" class="content">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-top: 1px solid #f3f4f6; padding-top: 32px; border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                  <tr>
                    <td align="center">
                      <div style="margin-bottom: 16px;">
                        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
                          <tr>
                            <td bgcolor="#d40511" width="30" height="4" style="font-size: 1px; line-height: 4px;">&nbsp;</td>
                            <td bgcolor="#ffcc00" width="30" height="4" style="font-size: 1px; line-height: 4px;">&nbsp;</td>
                          </tr>
                        </table>
                      </div>
                      <p style="margin: 0; font-size: 11px; color: #9ca3af; text-transform: uppercase; font-weight: 900; letter-spacing: 2px; font-family: Arial, sans-serif;">DHL Logistics • XML Validation System</p>
                      <p style="margin: 12px 0 0 0; font-size: 11px; color: #d1d5db; font-family: Arial, sans-serif;">© ${year} DHL. Todos os direitos reservados.</p>
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
