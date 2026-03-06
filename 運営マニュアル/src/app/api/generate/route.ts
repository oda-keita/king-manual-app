// APIルート: 複数ファイルアップロード → テキスト抽出 → Gemini API → 統合PowerPoint生成
import { NextRequest, NextResponse } from "next/server";
import { parseFile } from "@/lib/parseFile";
import { analyzeWithGemini, ContentData } from "@/lib/gemini";
import { generatePptx, FixedSlideAssets } from "@/lib/generatePptx";



/** エラーレスポンスを統一的に生成するヘルパー */
function errorResponse(
    message: string,
    errorType: "file_parse" | "ai_analysis" | "slide_generation" | "validation" | "unknown",
    status: number = 500,
    fileName?: string
) {
    console.error(`[API] エラー (${errorType}${fileName ? `, ${fileName}` : ""}):`, message);
    return NextResponse.json(
        { error: message, errorType, fileName },
        { status }
    );
}

export async function POST(request: NextRequest) {
    try {
        // 1. FormDataから複数ファイルとレイアウト画像を取得
        const formData = await request.formData();
        const files = formData.getAll("files") as File[];
        const layoutImages = formData.getAll("layoutImages") as string[];
        const timelineImageB64 = formData.get("timelineImage") as string | null;
        const roleImageB64 = formData.get("roleImage") as string | null;

        console.log(`[API] 設計書ファイル: ${files.length}件, レイアウト画像: ${layoutImages.length}ページ, TL画像: ${timelineImageB64 ? "あり" : "なし"}, 役職割画像: ${roleImageB64 ? "あり" : "なし"}`);

        if (!files || files.length === 0) {
            return errorResponse(
                "ファイルが選択されていません。",
                "validation",
                400
            );
        }

        // バリデーション
        for (const file of files) {
            if (file.size > 10 * 1024 * 1024) {
                return errorResponse(
                    `ファイル「${file.name}」のサイズが大きすぎます（上限: 10MB）。`,
                    "validation",
                    400,
                    file.name
                );
            }
            const ext = file.name.toLowerCase().split(".").pop();
            if (!ext || !["pdf", "docx"].includes(ext)) {
                return errorResponse(
                    `ファイル「${file.name}」は対応していません。対応形式は .pdf / .docx です。`,
                    "validation",
                    400,
                    file.name
                );
            }
        }

        // 2. 各ファイルを順番に処理（テキスト抽出 → Gemini API）
        const allContents: ContentData[] = [];

        for (const file of files) {
            console.log(`[API] 処理中: ${file.name}`);

            // --- テキスト抽出 ---
            let text: string;
            try {
                const arrayBuffer = await file.arrayBuffer();
                const buffer = Buffer.from(arrayBuffer);
                text = await parseFile(buffer, file.name);
            } catch (parseErr: unknown) {
                const msg = parseErr instanceof Error ? parseErr.message : String(parseErr);
                return errorResponse(
                    `ファイル読み取りエラー: ${msg}`,
                    "file_parse",
                    500,
                    file.name
                );
            }

            // --- Gemini APIで解析 ---
            let contents: ContentData[];
            try {
                contents = await analyzeWithGemini(text, layoutImages.length > 0 ? layoutImages : undefined);
            } catch (aiErr: unknown) {
                const msg = aiErr instanceof Error ? aiErr.message : String(aiErr);
                return errorResponse(
                    `AI解析エラー（${file.name}）: ${msg}`,
                    "ai_analysis",
                    500,
                    file.name
                );
            }

            allContents.push(...contents);

            console.log(
                `[API] ${file.name} → ${contents.length}コンテンツ抽出完了`
            );
        }

        if (allContents.length === 0) {
            return errorResponse(
                "アップロードされたファイルからコンテンツを抽出できませんでした。ファイルの内容を確認してください。",
                "ai_analysis"
            );
        }

        // 3. 全コンテンツを統合してPowerPoint生成
        console.log(`[API] 統合スライド生成中... (${allContents.length}コンテンツ)`);

        // 固定スライド用アセット
        const assets: FixedSlideAssets = {
            timelineImage: timelineImageB64 || undefined,
            roleImage: roleImageB64 || undefined,
        };

        let pptxBuffer: Buffer;
        try {
            pptxBuffer = await generatePptx(
                allContents,
                layoutImages.length > 0 ? layoutImages : undefined,
                assets
            );
        } catch (pptxErr: unknown) {
            const msg = pptxErr instanceof Error ? pptxErr.message : String(pptxErr);
            return errorResponse(
                `スライド生成エラー: ${msg}`,
                "slide_generation"
            );
        }

        // 4. バイナリレスポンスを返す
        return new NextResponse(new Uint8Array(pptxBuffer), {
            status: 200,
            headers: {
                "Content-Type":
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                "Content-Disposition": `attachment; filename="event-slides.pptx"`,
            },
        });
    } catch (error) {
        console.error("[API] 予期しないエラー:", error);
        const message =
            error instanceof Error
                ? error.message
                : "予期しないエラーが発生しました。";
        return errorResponse(message, "unknown");
    }
}
