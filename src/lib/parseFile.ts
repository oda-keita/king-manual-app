// PDF/Wordファイルからテキストを抽出するユーティリティ

import pdf from "pdf-parse";
import mammoth from "mammoth";

/**
 * ファイル名の拡張子に基づいてPDFまたはWordファイルからテキストを抽出
 */
export async function parseFile(
    buffer: Buffer,
    fileName: string
): Promise<string> {
    const ext = fileName.toLowerCase().split(".").pop();

    if (ext === "pdf") {
        return parsePdf(buffer, fileName);
    } else if (ext === "docx") {
        return parseDocx(buffer, fileName);
    } else {
        throw new Error(
            `サポートされていないファイル形式です: .${ext}\n対応形式: .pdf, .docx`
        );
    }
}

/** PDFからテキスト抽出 */
async function parsePdf(buffer: Buffer, fileName: string): Promise<string> {
    try {
        const data = await pdf(buffer);
        if (!data.text || data.text.trim().length === 0) {
            throw new Error(
                `「${fileName}」からテキストを抽出できませんでした。スキャンPDF（画像のみのPDF）の場合はOCR対応のPDFをお使いください。`
            );
        }
        return data.text;
    } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        // 既にカスタムメッセージで投げた場合はそのまま再throw
        if (msg.includes("テキストを抽出できませんでした")) {
            throw err;
        }
        // パスワード保護PDF
        if (
            msg.toLowerCase().includes("password") ||
            msg.toLowerCase().includes("encrypted")
        ) {
            throw new Error(
                `「${fileName}」はパスワード保護されているため読み取れません。パスワードを解除してから再度アップロードしてください。`
            );
        }
        // バッファ破損・不正なPDF
        if (
            msg.toLowerCase().includes("invalid") ||
            msg.toLowerCase().includes("corrupt") ||
            msg.toLowerCase().includes("not a pdf")
        ) {
            throw new Error(
                `「${fileName}」は破損しているか、正しいPDFファイルではありません。別のファイルをお試しください。`
            );
        }
        console.error(`[parseFile] PDF解析エラー (${fileName}):`, msg);
        throw new Error(
            `「${fileName}」のPDF解析中にエラーが発生しました: ${msg}`
        );
    }
}

/** Word(.docx)からテキスト抽出 */
async function parseDocx(buffer: Buffer, fileName: string): Promise<string> {
    try {
        const result = await mammoth.extractRawText({ buffer });
        if (!result.value || result.value.trim().length === 0) {
            throw new Error(
                `「${fileName}」からテキストを抽出できませんでした。ファイルが空か、対応していない形式の可能性があります。`
            );
        }
        return result.value;
    } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        if (msg.includes("テキストを抽出できませんでした")) {
            throw err;
        }
        // .doc形式の誤アップロード
        if (
            msg.toLowerCase().includes("could not find") ||
            msg.toLowerCase().includes("end of central directory")
        ) {
            throw new Error(
                `「${fileName}」は古い.doc形式か、破損した.docxファイルです。.docx形式で保存し直してからアップロードしてください。`
            );
        }
        console.error(`[parseFile] Word解析エラー (${fileName}):`, msg);
        throw new Error(
            `「${fileName}」のWord解析中にエラーが発生しました: ${msg}`
        );
    }
}
