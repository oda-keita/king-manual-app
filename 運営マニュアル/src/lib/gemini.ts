// Gemini APIを使ってイベント設計書を解析し、インフォグラフィック型スライド用JSONを生成

import { GoogleGenerativeAI } from "@google/generative-ai";

/** 階層付き箇条書きアイテム */
export interface BulletItem {
    text: string;
    indent_level: 0 | 1;
}

/** コンテンツ基本情報 */
export interface BasicInfo {
    content_name: string;
    date: string;
    time: string;
    location: string;
}

/** タイムライン項目 */
export interface TimelineItem {
    start_time: string;
    end_time: string;
    action: string;
}

/** 重要事項グループ */
export interface ImportantNote {
    display_type: "text_block" | "card_alert"; // text_block=テキストのみ, card_alert=カード強調
    card_theme?: "danger" | "info" | "warning"; // card_alert時のテーマ色
    heading: string;       // 絵文字付き見出し
    bullets: BulletItem[]; // 階層付き箇条書き（インラインタグ含む）
}

/** クリティカルアラート */
export interface CriticalAlert {
    action: string;
    bullets: BulletItem[]; // 階層付き箇条書き
}

/** コンテンツデータの型定義 */
export interface ContentData {
    basic_info: BasicInfo;
    timeline: TimelineItem[];
    important_notes: ImportantNote[];
    critical_timeline_alerts: CriticalAlert[];
    layout_page_number: number | null;
}

/**
 * Geminiのレスポンステキストから純粋なJSON文字列を抽出する
 */
function extractJsonFromResponse(responseText: string): string {
    let str = responseText.trim().replace(/^\uFEFF/, "");

    // Markdownコードブロック除去
    const codeBlockMatch = str.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (codeBlockMatch) str = codeBlockMatch[1].trim();

    // JSON配列/オブジェクト抽出
    if (!str.startsWith("[") && !str.startsWith("{")) {
        const arrayMatch = str.match(/(\[[\s\S]*\])/);
        if (arrayMatch) str = arrayMatch[1].trim();
        else {
            const objMatch = str.match(/(\{[\s\S]*\})/);
            if (objMatch) str = objMatch[1].trim();
        }
    }

    // 制御文字除去
    str = str.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "");
    return str;
}

/** Gemini APIにテキスト（+レイアウト画像）を送信し、コンテンツデータJSON配列を取得 */
export async function analyzeWithGemini(text: string, layoutImages?: string[]): Promise<ContentData[]> {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey || apiKey === "your_api_key_here") {
        throw new Error("GEMINI_API_KEYが設定されていません。.env.localにAPIキーを設定してください。");
    }

    const genAI = new GoogleGenerativeAI(apiKey);
    const model = genAI.getGenerativeModel({ model: "gemini-2.5-pro" });

    const systemPrompt = `あなたは厳格なイベント運営ディレクターです。
現場スタッフが1秒で理解して動けるよう、抽出する情報は**「絶対に長文を書かず、すべて短い箇条書き（体言止め、または簡潔な動詞終わり）」**にしてください。

【絶対禁止の表現】
- 「〜すること」「〜してください」「〜を徹底する」「〜に配慮した行動を徹底する」 → 禁止
- 「〜するように」「〜に注意する」 → 禁止
- 1つの箇条書きが40文字を超える → 禁止（分割せよ）

【必須の表現ルール】
- 体言止め: 「私語厳禁」「全員着席」「荷物は机の下」
- 短い動詞: 「前方へ移動」「#報告チャンネルで共有」「フォーム提出（15:00厳守）」
- 主語明示: 「TSは〇〇」「参加者には〇〇させる」

必ずJSON配列として出力してください（コンテンツが1つでも配列で返すこと）。

【複数コンテンツの全数抽出（最重要）】
PDFに複数のコンテンツ（例: Day0, Day1, Day2, Day3、または開会式・コンサル・予選審査・決勝審査・閉会式等）が含まれている場合は、その全てを個別のオブジェクトとしてJSON配列に含めてください。1つだけ出力して残りを省略することは絶対に禁止です。設計書に記載された全コンテンツを残らず抽出してください。
Markdown（\`\`\`jsonなど）は絶対に含めないでください。

【インラインタグ（テキスト内の強調装飾）】
箇条書きのテキスト内で強調すべき箇所には、以下のタグを埋め込んでください。
1. <bold_blue>重要キーワード</bold_blue> ：重要だが警告ではないキーワード（青色強調）
2. <bold_red>警告キーワード</bold_red> ：絶対にミスが許されないキーワード（赤色警告）
3. <marker_yellow>注目キーワード</marker_yellow> ：特に視線を誘導したい箇所（黄色マーカー）
※全ての箇条書きにタグを付ける必要はなし。本当に強調すべき重要箇所のみに使用。

【display_typeの使い分け】
important_notesの各グループには "display_type" を指定してください。
- "text_block" ：通常の注意事項。カード（箱）にせず、テキストのみでスッキリ見せる。
- "card_alert" ：致命的なルールの塊。背景色付きカードにして目立たせる。card_theme（danger/info/warning）も指定。
全てをcard_alertにしない。本当に重要な塊のみcard_alertにし、残りはtext_blockにすること。

【出力JSON構造】
[
  {
    "basic_info": {
      "content_name": "コンテンツ名（例: Spring Business Contest 2026〈決勝審査〉）",
      "date": "日程（例: 3月5日）",
      "time": "全体の開始〜終了時刻（例: 15:05〜17:00）",
      "location": "階と部屋番号のみ（例: 3F 301、B1F 多目的室）"
    },
    "timeline": [
      { "start_time": "10:00", "end_time": "10:30", "action": "アクション（簡潔に）" }
    ],
    "important_notes": [
      {
        "display_type": "text_block",
        "heading": "🏆 表彰順",
        "bullets": [
          { "text": "オーディエンス賞 → 三位 → 二位 → 一位", "indent_level": 0 },
          { "text": "呼ばれたら前方移動 → 賞状読み上げ → 賞品手渡し", "indent_level": 1 }
        ]
      },
      {
        "display_type": "card_alert",
        "card_theme": "danger",
        "heading": "⚠️ 厳守事項",
        "bullets": [
          { "text": "<bold_red>参加者の前での私語は厳禁</bold_red>", "indent_level": 0 },
          { "text": "遅刻者は<marker_yellow>#105_staff_連絡チャンネル</marker_yellow>で報告", "indent_level": 0 }
        ]
      }
    ],
    "critical_timeline_alerts": [
      {
        "action": "審査時間中",
        "bullets": [
          { "text": "<bold_red>参加者を静粛にさせる</bold_red>", "indent_level": 0 },
          { "text": "私語厳禁・足音配慮", "indent_level": 1 }
        ]
      }
    ],
    "layout_page_number": null
  }
]

【重要ルール】
- basic_info.date, basic_info.time, basic_info.location は設計書から必ず抽出。不明なら空文字。
- timeline は設計書の全スケジュール項目を網羅。件数制限なし。
- important_notes はカテゴリ別に抽出（表彰順・提出ルール・禁止事項・連絡フロー・役割分担等）。件数制限なし。
- critical_timeline_alerts は「ミスしたら取り返しがつかないアクション」を全て抽出。件数制限なし。
- bullets の各要素は { "text": "短い箇条書き", "indent_level": 0 } 形式。indent_level は 0（メイン）または 1（サブ詳細）。
- 必ず配列形式（[{...}]）のJSONのみを出力。

【Winterマニュアル水準の思考・補完プロセス（最重要）】
入力される「設計書」のテキストは計画段階のものであり、現場マニュアルとしては情報が抽象的です。以下のルールで「現場レベルのアクション」に変換・補完してから出力してください。

1. 「物理的動線」の具体化：
「移動」「設営」→「荷物は机の下」「〇〇の指示に従い移動」等、現場で迷わないレベルに。
2. 「ペナルティとルール」の厳格化：
「1分遅れで1点減点」「未提出は0点」「公式垢使用禁止」等、シビアな基準を明文化。
3. 「連絡フロー」の特定：
「#〇〇チャンネルにて報告」「指定フォームで提出」等、具体的なツール・チャンネル名を補完。
4. 「主語（誰が）」の明確な切り分け：
「TSは〇〇を行う」「参加者には〇〇させる」と主語を絶対に混同しない。
5. 「懸念点」の警告化：
設計書の懸念点を「🚨 重要アラート」として具体的行動指示に変換。
6. 「運営スタッフとしての行動規範」の注入：
「私語厳禁」「足音配慮」「社会人には道を譲る」「ネガティブ発言禁止」等をアラートに組込。
7. 「空間認識・立ち位置」の指定：
「スクリーン前横切り禁止」「後方遠回りで移動」「社会人の視界外に立つ」等。
8. 「情報管理・参加者との境界線」：
「Slack画面を参加者に見せない」「機密情報口外禁止」等。
9. 「イレギュラー対応と待機姿勢」：
「勝手に外出しない」「居場所を常に共有」「待機中もすぐ動ける姿勢」等。

【レイアウト画像の紐づけ】
テキストデータと共に、レイアウト図の画像データを複数枚送信している場合があります。
各画像の中に書かれているタイトル（「開会式」「プランニング」など）を視覚的に読み取り、現在のコンテンツと一致するレイアウト画像がある場合は、その「画像のページ番号（1始まりのインデックス）」をJSONの layout_page_number として出力してください。該当する画像がない場合は null を指定してください。レイアウト画像が一枚も送信されていない場合は、layout_page_number は常に null としてください。`;

    let responseText: string;
    try {
        // マルチモーダル送信用のpartsを構築
        const parts: Array<{ text: string } | { inlineData: { mimeType: string; data: string } }> = [
            { text: systemPrompt },
            { text: `以下がイベント設計書のテキストです:\n\n${text}` },
        ];

        // レイアウト画像があれば inlineData として追加
        if (layoutImages && layoutImages.length > 0) {
            console.log(`[Gemini] レイアウト画像 ${layoutImages.length}枚をマルチモーダル送信`);
            parts.push({ text: `\n以下にレイアウト図の画像を${layoutImages.length}枚送信します。各画像のページ番号は1始まりです。` });
            for (const img of layoutImages) {
                parts.push({
                    inlineData: {
                        mimeType: "image/png",
                        data: img,
                    },
                });
            }
        }

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const result = await model.generateContent(parts as any);
        responseText = result.response.text();
    } catch (apiError: unknown) {
        const errMsg = apiError instanceof Error ? apiError.message : String(apiError);
        console.error("Gemini API呼び出しエラー:", errMsg);
        if (errMsg.includes("429") || errMsg.toLowerCase().includes("rate limit") || errMsg.toLowerCase().includes("quota")) {
            throw new Error("Gemini APIのレートリミットに達しました（429エラー）。しばらく時間をおいてから再度お試しください。");
        }
        if (errMsg.includes("403") || errMsg.toLowerCase().includes("forbidden")) {
            throw new Error("Gemini APIへのアクセスが拒否されました（403エラー）。APIキーが正しいか確認してください。");
        }
        if (errMsg.includes("401") || errMsg.toLowerCase().includes("unauthorized")) {
            throw new Error("Gemini APIの認証に失敗しました（401エラー）。APIキーが有効か確認してください。");
        }
        throw new Error(`Gemini API呼び出しに失敗しました: ${errMsg}`);
    }

    console.log("[Gemini] レスポンス先頭200文字:", responseText.substring(0, 200));
    const jsonStr = extractJsonFromResponse(responseText);

    try {
        let parsed = JSON.parse(jsonStr);

        // オブジェクト単体で返ってきた場合は配列に変換
        if (!Array.isArray(parsed)) {
            parsed = [parsed];
        }

        const contents: ContentData[] = parsed;

        for (const content of contents) {
            // basic_info のサニタイズ
            if (!content.basic_info || typeof content.basic_info !== "object") {
                content.basic_info = { content_name: "名称未設定", date: "", time: "", location: "" };
            }
            content.basic_info.content_name = String(content.basic_info.content_name || "名称未設定");
            content.basic_info.date = String(content.basic_info.date || "");
            content.basic_info.time = String(content.basic_info.time || "");
            content.basic_info.location = String(content.basic_info.location || "");

            // timeline のサニタイズ
            if (!Array.isArray(content.timeline)) content.timeline = [];
            content.timeline = content.timeline.map((item) => ({
                start_time: String(item?.start_time || ""),
                end_time: String(item?.end_time || ""),
                action: String(item?.action || ""),
            })).filter(item => item.action.trim().length > 0);

            // important_notes のサニタイズ（BulletItem + display_type対応）
            if (!Array.isArray(content.important_notes)) content.important_notes = [];
            content.important_notes = content.important_notes.map((note) => {
                // display_type のサニタイズ
                const dt = String(note?.display_type || "text_block");
                const display_type: "text_block" | "card_alert" = dt === "card_alert" ? "card_alert" : "text_block";
                // card_theme のサニタイズ
                const ct = String(note?.card_theme || "");
                const card_theme: "danger" | "info" | "warning" | undefined =
                    ct === "danger" ? "danger" : ct === "info" ? "info" : ct === "warning" ? "warning" : undefined;
                return {
                    display_type,
                    card_theme,
                    heading: String(note?.heading || ""),
                    bullets: sanitizeBullets(note?.bullets),
                };
            }).filter(n => n.heading.trim().length > 0 && n.bullets.length > 0);

            // critical_timeline_alerts のサニタイズ（BulletItem対応）
            if (!Array.isArray(content.critical_timeline_alerts)) content.critical_timeline_alerts = [];
            content.critical_timeline_alerts = content.critical_timeline_alerts.map((alert) => ({
                action: String(alert?.action || ""),
                bullets: sanitizeBullets(alert?.bullets),
            })).filter(a => a.action.trim().length > 0 && a.bullets.length > 0);

            // layout_page_number のサニタイズ
            if (content.layout_page_number !== null && content.layout_page_number !== undefined) {
                const pageNum = Number(content.layout_page_number);
                content.layout_page_number = isNaN(pageNum) ? null : pageNum;
            } else {
                content.layout_page_number = null;
            }
        }

        return contents;
    } catch (e) {
        console.error("[Gemini] レスポンス解析失敗。生テキスト:", responseText);
        console.error("[Gemini] クレンジング後テキスト:", jsonStr);
        throw new Error(
            `AIのレスポンスを解析できませんでした。もう一度お試しください。\n（技術詳細: ${e instanceof Error ? e.message : String(e)}）`
        );
    }
}

/**
 * bullets配列をBulletItem[]にサニタイズするヘルパー
 * - 旧形式（string[]）にも対応: string → { text: string, indent_level: 0 }
 * - オブジェクト形式: { text, indent_level } をそのまま変換
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function sanitizeBullets(bullets: any): BulletItem[] {
    if (!Array.isArray(bullets)) return [];
    return bullets
        .map((b: unknown) => {
            if (typeof b === "string") {
                // 旧形式: string → BulletItem
                return { text: b.trim(), indent_level: 0 as const };
            }
            if (b && typeof b === "object" && "text" in b) {
                const obj = b as { text: unknown; indent_level?: unknown };
                const text = String(obj.text || "").trim();
                const indent = Number(obj.indent_level) === 1 ? 1 as const : 0 as const;
                return { text, indent_level: indent };
            }
            return null;
        })
        .filter((b: BulletItem | null): b is BulletItem => b !== null && b.text.length > 0);
}
