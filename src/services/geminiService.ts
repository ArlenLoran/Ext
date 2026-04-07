import { GoogleGenAI } from "@google/genai";

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

export async function extractInvoiceNumber(pdfBlob: Blob): Promise<string | null> {
  try {
    // Convert blob to base64
    const reader = new FileReader();
    const base64Promise = new Promise<string>((resolve, reject) => {
      reader.onloadend = () => {
        const result = reader.result as string;
        const base64 = result.split(',')[1];
        resolve(base64);
      };
      reader.onerror = reject;
    });
    reader.readAsDataURL(pdfBlob);
    const base64Data = await base64Promise;

    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: [
        {
          inlineData: {
            mimeType: "application/pdf",
            data: base64Data,
          },
        },
        {
          text: "Extract the Invoice Number (Número da Nota Fiscal or nNF) from this PDF. Return ONLY the digits of the number. If not found, return 'Não encontrado'.",
        },
      ],
    });

    const text = response.text?.trim() || '';
    if (text.toLowerCase().includes('não encontrado')) return null;
    
    // Clean the number (remove non-digits)
    const cleaned = text.replace(/\D/g, '');
    return cleaned || null;
  } catch (err) {
    console.error('Erro ao extrair número da NF com Gemini:', err);
    return null;
  }
}
