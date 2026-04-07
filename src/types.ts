export interface PdfFile {
  name: string;
  serverRelativeUrl: string;
  timeCreated: string;
  size: number;
  invoiceNumber?: string;
  isIndexing?: boolean;
}
