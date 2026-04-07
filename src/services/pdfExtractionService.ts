import * as pdfjsLib from 'pdfjs-dist';

// Configure worker using jsdelivr which is highly reliable
// Note: pdfjs-dist 4.0+ uses .mjs for the worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.mjs`;

export async function extractInvoiceNumberWithoutAI(pdfBlob: Blob): Promise<string | null> {
  try {
    const arrayBuffer = await pdfBlob.arrayBuffer();
    const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
    const pdf = await loadingTask.promise;
    
    let fullText = '';
    
    // Read first 2 pages (usually enough for NF)
    const numPages = Math.min(pdf.numPages, 2);
    
    for (let i = 1; i <= numPages; i++) {
      const page = await pdf.getPage(i);
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map((item: any) => item.str).join(' ');
      fullText += pageText + '\n';
    }

    // Specific patterns for DACE (Declaração Auxiliar de Conteúdo Eletrônica)
    const patterns = [
      /Nº\s*(\d+[\.?\d+]*)/i, // Matches "Nº 7.857"
      /OS\s*[:\-]?\s*(\d+)/i, // Matches "OS: 62981129"
      /nNF\s*[:\-]?\s*(\d+[\.?\d+]*)/i,
      /Número\s*da\s*Nota\s*[:\-]?\s*(\d+[\.?\d+]*)/i,
      /NF-e\s*nº\s*(\d+[\.?\d+]*)/i,
      /(\d{3}\.\d{3}\.\d{3})/
    ];

    let invoiceNumber = null;
    let osNumber = null;

    for (const pattern of patterns) {
      const match = fullText.match(pattern);
      if (match && match[1]) {
        const cleaned = match[1].replace(/\D/g, '');
        const value = cleaned.replace(/^0+/, '') || cleaned;
        
        if (pattern.source.includes('OS')) {
          osNumber = value;
        } else if (!invoiceNumber) {
          invoiceNumber = value;
        }
      }
    }

    // If we found an OS but no DACE number, or if we want to combine them
    // For now, let's return the DACE number as primary, or OS as fallback
    return invoiceNumber || osNumber;

    return null;
  } catch (err) {
    console.error('Erro ao extrair texto do PDF:', err);
    return null;
  }
}
