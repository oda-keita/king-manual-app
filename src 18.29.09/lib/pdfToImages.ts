// ブラウザ上でpdfjs-distを使い、PDFの各ページをBase64 PNG画像に変換する
// サーバーサイドではDOMMatrix等が未定義のため、動的importでクライアント側のみで読み込む

/**
 * PDFファイルを読み込み、各ページをBase64 PNG文字列の配列に変換する
 * @param file アップロードされたPDFファイル
 * @param scale レンダリングスケール（デフォルト2.0で高解像度）
 * @returns Base64 PNG文字列の配列（"data:image/png;base64,..." を含まない生Base64）
 */
export async function pdfToBase64Images(
    file: File,
    scale: number = 2.0
): Promise<string[]> {
    // pdfjs-distを動的にインポート（SSR時のDOMMatrixエラーを回避）
    const pdfjsLib = await import("pdfjs-dist");

    // PDF.jsワーカーの設定（publicディレクトリから静的配信）
    if (!pdfjsLib.GlobalWorkerOptions.workerSrc) {
        pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";
    }

    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    const images: string[] = [];

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
        const page = await pdf.getPage(pageNum);
        const viewport = page.getViewport({ scale });

        // Canvas要素を作成してPDFページをレンダリング
        const canvas = document.createElement("canvas");
        canvas.width = viewport.width;
        canvas.height = viewport.height;

        const context = canvas.getContext("2d");
        if (!context) {
            throw new Error(`Canvas 2Dコンテキストの取得に失敗しました（ページ ${pageNum}）`);
        }

        await page.render({
            canvasContext: context,
            viewport: viewport,
            canvas: canvas,
        }).promise;

        // Canvas → Base64 PNG（プレフィックス "data:image/png;base64," を除去した生データ）
        const dataUrl = canvas.toDataURL("image/png");
        const base64 = dataUrl.replace(/^data:image\/png;base64,/, "");
        images.push(base64);

        // メモリ解放
        page.cleanup();
    }

    return images;
}
