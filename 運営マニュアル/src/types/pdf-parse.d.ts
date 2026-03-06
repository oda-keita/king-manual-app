// pdf-parseの型定義
declare module "pdf-parse" {
    interface PDFData {
        numpages: number;
        numrender: number;
        info: Record<string, unknown>;
        metadata: Record<string, unknown>;
        version: string;
        text: string;
    }

    function pdf(dataBuffer: Buffer): Promise<PDFData>;
    export default pdf;
}

// pdfjs-distワーカーの型定義
declare module "pdfjs-dist/build/pdf.worker.min.mjs" {
    const workerSrc: string;
    export default workerSrc;
}
