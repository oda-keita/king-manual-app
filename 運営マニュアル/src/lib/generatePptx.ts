// PptxGenJSを使って縦型PowerPointを生成
// デザイン: Miraial風IRモダンデザイン + 自動改ページ（オートページング）対応Winterマニュアル仕様
// 階層付き箇条書き・レイアウト画像の動的配置・最大3枚制限

import PptxGenJS from "pptxgenjs";
import type { ContentData, BasicInfo, ImportantNote, CriticalAlert, BulletItem } from "./gemini";

// ===== サイズ =====
const W = 7.5;
const H = 10;

// ===== 自動改ページ制御定数 =====
const PAGE_BOTTOM = 9.2;        // スライド下限Y座標
const CONTINUATION_TOP = 1.6;   // 2枚目以降のコンテンツ開始Y
const MIN_BODY_FONT = 11;       // 本文最小フォントサイズ（pt）
const MAX_SLIDES_PER_CONTENT = 3; // 1コンテンツあたりの最大スライド数

// ===== カラーパレット（セマンティックデザイン）=====
const C = {
    // セマンティックカラー（意味的な使い分け）
    baseDark: "003F8E",     // 見出し、重要テキスト、強調ライン
    baseBlue: "005BAC",     // 青の強調文字
    textMain: "1F2937",     // 基本の本文テキスト
    textSub: "64748B",      // 補足説明、表ヘッダーテキスト
    accentRed: "EF4444",    // 警告、致命的なルールの強調
    markerYellow: "FEF08A", // 背景ハイライト用マーカー
    bgCardDanger: "FEF2F2", // 警告カードの極薄い赤背景
    bgCardInfo: "F0F9FF",   // 補足カードの極薄い青背景
    bgCardWarning: "FFFBEB", // 警告カードの極薄い黄背景
    bgTableHead: "F8FAFC",  // 表のヘッダー背景

    // アクセントカラー
    amber: "F59E0B",
    amberSoft: "FEF3C7",
    amberDeep: "D97706",
    green: "22C55E",
    greenSoft: "D1FAE5",
    purple: "7C3AED",
    purpleSoft: "EDE9FE",
    teal: "0D9488",
    tealSoft: "CCFBF1",

    // 背景・グレー系
    white: "FFFFFF",
    bgColor: "FFFFFF",     // スライド全体の背景色（白ベース）
    borderLight: "E2E8F0", // 表の罫線用極薄グレー
    gray200: "E5E7EB",
    gray300: "D1D5DB",
    gray400: "9CA3AF",
};

const FONT = "Yu Gothic";        // ベースフォント（游ゴシック）
const FONT_NUM = "Oswald";       // 数字・英字強調用サンセリフ体

// ===== 固定スライド用アセット =====
export interface FixedSlideAssets {
    timelineImage?: string;  // 全体TL用画像Base64
    roleImage?: string;      // 役職割用画像Base64
}
const MX = 0.55; // 左右マージン
const CW = W - MX * 2;
const HEADER_Y = 0.06; // 上部アクセントライン太さ
const CONTENT_TOP = 1.6; // コンテンツ開始Y
const FOOTER_H = 0.3;

// ===== セクション見出しのアクセント色パレット =====
const SECTION_ACCENTS = [
    C.baseDark,
    C.green,
    C.purple,
    C.amberDeep,
    C.teal,
];

// ===== card_alert テーマ色マッピング（極薄い背景色） =====
const ALERT_THEMES = {
    danger: { bg: C.bgCardDanger, accent: C.accentRed, text: C.accentRed },
    info: { bg: C.bgCardInfo, accent: C.baseDark, text: C.baseDark },
    warning: { bg: C.bgCardWarning, accent: C.amberDeep, text: C.amberDeep },
};

// ===== インラインタグパーサー =====
// <bold_blue>...</bold_blue>, <bold_red>...</bold_red>, <marker_yellow>...</marker_yellow>
// をPptxGenJSのテキストフォーマット配列に変換するヘルパー
type TextPart = { text: string; options: { fontSize: number; fontFace: string; color: string; bold?: boolean; highlight?: string } };

function parseInlineTags(rawText: string, baseFontSize: number, baseColor: string): TextPart[] {
    const tagPattern = /<(bold_blue|bold_red|marker_yellow)>(.*?)<\/\1>/g;
    const parts: TextPart[] = [];
    let lastIndex = 0;
    let match: RegExpExecArray | null;

    while ((match = tagPattern.exec(rawText)) !== null) {
        // タグの前のプレーンテキスト
        if (match.index > lastIndex) {
            parts.push({
                text: rawText.slice(lastIndex, match.index),
                options: { fontSize: baseFontSize, fontFace: FONT, color: baseColor },
            });
        }
        const tagType = match[1];
        const tagContent = match[2];
        if (tagType === "bold_blue") {
            parts.push({
                text: tagContent,
                options: { fontSize: baseFontSize + 1, fontFace: FONT, color: C.baseBlue, bold: true },
            });
        } else if (tagType === "bold_red") {
            parts.push({
                text: tagContent,
                options: { fontSize: baseFontSize + 1, fontFace: FONT, color: C.accentRed, bold: true },
            });
        } else if (tagType === "marker_yellow") {
            parts.push({
                text: tagContent,
                options: { fontSize: baseFontSize, fontFace: FONT, color: C.textMain, bold: true, highlight: C.markerYellow },
            });
        }
        lastIndex = match.index + match[0].length;
    }
    // 残りのプレーンテキスト
    if (lastIndex < rawText.length) {
        parts.push({
            text: rawText.slice(lastIndex),
            options: { fontSize: baseFontSize, fontFace: FONT, color: baseColor },
        });
    }
    // タグが1つもなければ元テキストをそのまま返す
    if (parts.length === 0) {
        parts.push({
            text: rawText,
            options: { fontSize: baseFontSize, fontFace: FONT, color: baseColor },
        });
    }
    return parts;
}

// ===== スライド生成カウンタ =====
interface SlideContext {
    slide: PptxGenJS.Slide;
    y: number;
    slideCount: number; // このコンテンツで使用したスライド数
}

// ============================================================
//  メインエクスポート
// ============================================================
export async function generatePptx(
    allContents: ContentData[],
    layoutImages?: string[],
    assets?: FixedSlideAssets
): Promise<Buffer> {
    try {
        const pptx = new PptxGenJS();
        pptx.defineLayout({ name: "PORTRAIT", width: W, height: H });
        pptx.layout = "PORTRAIT";
        pptx.author = "Event Slide Generator";
        pptx.subject = "イベント運営マニュアル";

        // ===== 固定スライド =====
        buildFixedTitleSlide(pptx);                     // スライド1: タイトル
        buildTocSlide(pptx, allContents);               // スライド2: 目次
        buildTimelineImageSlide(pptx, assets?.timelineImage); // スライド3: 全体TL
        buildRoleImageSlide(pptx, assets?.roleImage);   // スライド4: 役職割
        buildDressCodeAndSubmissionsSlide(pptx);        // スライド5: 服装+外出+提出物
        buildHearingRulesSlide(pptx);                   // スライド6: ヒアリング注意事項
        // Slackチャンネルガイド（4枚に統合）
        buildSlackCombinedPublicAnd100Slide(pptx);   // パブリック+100番台
        buildSlackCombined200to500Slide(pptx);        // 200〜500番台
        buildSlackHealthSlide(pptx);                  // 体調管理
        buildSlackAttendanceSlide(pptx);

        // ===== 動的コンテンツスライド =====
        for (let i = 0; i < allContents.length; i++) {
            buildContentSlides(pptx, allContents[i], i + 1, allContents.length, layoutImages);
        }

        const output = await pptx.write({ outputType: "nodebuffer" });
        return output as Buffer;
    } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        console.error("[generatePptx] エラー:", msg);
        throw new Error(`スライド生成中にエラーが発生しました。\n（${msg}）`);
    }
}

// ============================================================
//  固定スライド1: タイトルスライド（BaseDark背景 + 白文字）
// ============================================================
function buildFixedTitleSlide(pptx: PptxGenJS) {
    const slide = pptx.addSlide();
    slide.background = { color: C.baseDark };

    // 上部アクセントライン
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: W, h: 0.1,
        fill: { color: C.baseBlue },
    });

    // メインタイトル
    slide.addText("Spring Business\nContest 2026", {
        x: MX, y: 2.5, w: CW, h: 2.5,
        fontSize: 38, fontFace: FONT, color: C.white, bold: true,
        lineSpacingMultiple: 1.2,
    });

    // サブタイトル
    slide.addText("運営マニュアル", {
        x: MX, y: 5.2, w: CW, h: 0.8,
        fontSize: 24, fontFace: FONT, color: C.white, bold: true,
    });

    // 下部アクセントライン
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 6.2, w: 2.0, h: 0.05,
        fill: { color: C.baseBlue },
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX + 2.1, y: 6.2, w: CW - 2.1, h: 0.015,
        fill: { color: "1A5BB0" },
    });

    // 日付等
    slide.addText("EVENT OPERATIONS MANUAL", {
        x: MX, y: 6.5, w: CW, h: 0.4,
        fontSize: 10, fontFace: FONT_NUM, color: "7BABD4", bold: true,
    });
}

// ============================================================
//  固定スライド2: 目次（TOC）
// ============================================================
function buildTocSlide(pptx: PptxGenJS, allContents: ContentData[]) {
    const slide = pptx.addSlide();
    slide.background = { color: C.bgColor };

    // 上バー
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: W, h: HEADER_Y,
        fill: { color: C.baseDark },
    });

    // 見出し
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.4, w: 0.07, h: 0.4,
        fill: { color: C.baseDark },
    });
    slide.addText("目次", {
        x: MX + 0.18, y: 0.4, w: 3, h: 0.4,
        fontSize: 24, fontFace: FONT, color: C.textMain, bold: true,
        valign: "middle",
    });

    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.95, w: CW, h: 0.015,
        fill: { color: C.gray200 },
    });

    // 固定セクション（全セクションを網羅）
    const fixedItems = [
        "全体タイムライン",
        "役職割",
        "服装・外出可能時間 & 提出物一覧",
        "ヒアリング実施時の注意事項",
        "Slack: パブリック & 100番台",
        "Slack: 200〜500番台",
        "Slack: 体調管理の使い方",
        "Slack: 欠席早退遅刻連絡の使い方",
    ];

    let y = 1.3;
    const itemH = 0.30;

    // 固定項目
    fixedItems.forEach((name, i) => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: MX, y, w: 0.42, h: itemH,
            fill: { color: C.baseDark }, rectRadius: 0.05,
        });
        slide.addText(`${String(i + 1).padStart(2, "0")}`, {
            x: MX, y, w: 0.42, h: itemH,
            fontSize: 12, fontFace: FONT_NUM, color: C.white, bold: true,
            align: "center", valign: "middle",
        });
        slide.addText(name, {
            x: MX + 0.52, y, w: CW - 0.52, h: itemH,
            fontSize: 12, fontFace: FONT, color: C.textMain, bold: true,
            valign: "middle",
        });
        y += itemH + 0.06;
    });

    // 区切り
    y += 0.08;
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y, w: CW, h: 0.01,
        fill: { color: C.gray200 },
    });
    y += 0.15;

    slide.addText("各コンテンツ（Gemini抽出）", {
        x: MX, y, w: CW, h: 0.3,
        fontSize: 10, fontFace: FONT, color: C.textSub, bold: true,
    });
    y += 0.35;

    // 動的コンテンツ一覧
    allContents.forEach((c, i) => {
        const num = fixedItems.length + i + 1;
        slide.addShape(pptx.ShapeType.roundRect, {
            x: MX, y, w: 0.42, h: itemH,
            fill: { color: C.baseBlue }, rectRadius: 0.05,
        });
        slide.addText(`${String(num).padStart(2, "0")}`, {
            x: MX, y, w: 0.42, h: itemH,
            fontSize: 12, fontFace: FONT_NUM, color: C.white, bold: true,
            align: "center", valign: "middle",
        });

        const name = c.basic_info?.content_name || "名称未設定";
        const time = c.basic_info?.time || "";
        slide.addText(name + (time ? `  （${time}）` : ""), {
            x: MX + 0.52, y, w: CW - 0.52, h: itemH,
            fontSize: 11, fontFace: FONT, color: C.textMain,
            valign: "middle",
        });
        y += itemH + 0.04;
    });

    addFooter(pptx, slide);
}

// ============================================================
//  固定スライド3: 全体タイムライン画像
// ============================================================
function buildTimelineImageSlide(pptx: PptxGenJS, base64Image?: string) {
    const slide = pptx.addSlide();
    slide.background = { color: C.bgColor };

    // ヘッダー
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: W, h: HEADER_Y,
        fill: { color: C.baseDark },
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.35, w: 0.07, h: 0.4,
        fill: { color: C.baseDark },
    });
    slide.addText("全体タイムライン", {
        x: MX + 0.18, y: 0.3, w: CW - 1.5, h: 0.5,
        fontSize: 20, fontFace: FONT, color: C.textMain, bold: true,
        valign: "middle",
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.95, w: CW, h: 0.015,
        fill: { color: C.gray200 },
    });

    if (base64Image) {
        // 画像を最大限のサイズで配置
        const imgTop = 1.2;
        const imgBottom = PAGE_BOTTOM - 0.2;
        const availH = imgBottom - imgTop;
        const assumedAspect = 4 / 3;
        let drawW: number, drawH: number;

        if (CW / availH > assumedAspect) {
            drawH = availH;
            drawW = drawH * assumedAspect;
        } else {
            drawW = CW;
            drawH = drawW / assumedAspect;
        }

        const drawX = MX + (CW - drawW) / 2;
        const drawY = imgTop + (availH - drawH) / 2;

        slide.addImage({
            data: `image/png;base64,${base64Image}`,
            x: drawX, y: drawY, w: drawW, h: drawH,
        });
    } else {
        slide.addText("※ 全体TL用画像がアップロードされていません", {
            x: MX, y: 4.0, w: CW, h: 1.0,
            fontSize: 14, fontFace: FONT, color: C.textSub,
            align: "center", valign: "middle",
        });
    }

    addFooter(pptx, slide);
}

// ============================================================
//  固定スライド4: 役職割画像
// ============================================================
function buildRoleImageSlide(pptx: PptxGenJS, base64Image?: string) {
    const slide = pptx.addSlide();
    slide.background = { color: C.bgColor };

    // ヘッダー
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: W, h: HEADER_Y,
        fill: { color: C.baseDark },
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.35, w: 0.07, h: 0.4,
        fill: { color: C.baseDark },
    });
    slide.addText("役職割", {
        x: MX + 0.18, y: 0.3, w: CW - 1.5, h: 0.5,
        fontSize: 20, fontFace: FONT, color: C.textMain, bold: true,
        valign: "middle",
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.95, w: CW, h: 0.015,
        fill: { color: C.gray200 },
    });

    if (base64Image) {
        const imgTop = 1.2;
        const imgBottom = PAGE_BOTTOM - 0.2;
        const availH = imgBottom - imgTop;
        const assumedAspect = 4 / 3;
        let drawW: number, drawH: number;

        if (CW / availH > assumedAspect) {
            drawH = availH;
            drawW = drawH * assumedAspect;
        } else {
            drawW = CW;
            drawH = drawW / assumedAspect;
        }

        const drawX = MX + (CW - drawW) / 2;
        const drawY = imgTop + (availH - drawH) / 2;

        slide.addImage({
            data: `image/png;base64,${base64Image}`,
            x: drawX, y: drawY, w: drawW, h: drawH,
        });
    } else {
        slide.addText("※ 役職割用画像がアップロードされていません", {
            x: MX, y: 4.0, w: CW, h: 1.0,
            fontSize: 14, fontFace: FONT, color: C.textSub,
            align: "center", valign: "middle",
        });
    }

    addFooter(pptx, slide);
}

// ============================================================
//  固定スライド5: 服装・外出可能時間（上部）& 提出物一覧（下部）
// ============================================================
function buildDressCodeAndSubmissionsSlide(pptx: PptxGenJS) {
    const slide = pptx.addSlide();
    slide.background = { color: C.bgColor };

    // ヘッダー
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: W, h: HEADER_Y,
        fill: { color: C.baseDark },
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.35, w: 0.07, h: 0.4,
        fill: { color: C.baseDark },
    });
    slide.addText("服装・外出可能時間 & 提出物一覧", {
        x: MX + 0.18, y: 0.3, w: CW - 0.5, h: 0.5,
        fontSize: 18, fontFace: FONT, color: C.textMain, bold: true,
        valign: "middle",
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.95, w: CW, h: 0.015,
        fill: { color: C.gray200 },
    });

    // ===== 上部: 服装（左）+ 外出可能時間（右）2カラム =====
    const colW = (CW - 0.20) / 2;
    const leftX = MX;
    const rightX = MX + colW + 0.20;

    // --- 左: 服装 ---
    let y = 1.15;
    slide.addShape(pptx.ShapeType.rect, {
        x: leftX, y, w: 0.05, h: 0.25,
        fill: { color: C.baseDark },
    });
    slide.addText("👔 服装", {
        x: leftX + 0.12, y, w: colW - 0.12, h: 0.25,
        fontSize: 13, fontFace: FONT, color: C.baseDark, bold: true,
        valign: "middle",
    });
    y += 0.32;

    const dressItems = [
        { role: "運営スタッフ", dress: "オフィスカジュアル" },
        { role: "登壇者", dress: "オフィスカジュアル" },
        { role: "社対", dress: "スーツ" },
        { role: "カメラマン", dress: "色味抑えめの服" },
        { role: "TS", dress: "自由" },
        { role: "参加者", dress: "自由" },
        { role: "巡回", dress: "自由" },
    ];

    dressItems.forEach((item) => {
        slide.addText([
            { text: `• ${item.role}：`, options: { fontSize: 10, fontFace: FONT, color: C.textMain, bold: true } },
            { text: item.dress, options: { fontSize: 10, fontFace: FONT, color: C.textSub } },
        ], {
            x: leftX + 0.08, y, w: colW - 0.08, h: 0.24,
            valign: "middle",
        });
        y += 0.26;
    });

    // --- 右: 外出可能時間 ---
    let ry = 1.15;
    slide.addShape(pptx.ShapeType.rect, {
        x: rightX, y: ry, w: 0.05, h: 0.25,
        fill: { color: C.baseBlue },
    });
    slide.addText("🕐 外出可能時間", {
        x: rightX + 0.12, y: ry, w: colW - 0.12, h: 0.25,
        fontSize: 13, fontFace: FONT, color: C.baseDark, bold: true,
        valign: "middle",
    });
    ry += 0.38;

    const outingData = [
        { day: "Day 1", time: "11:00 - 20:00", note: "" },
        { day: "Day 2", time: "11:00~13:00 / 17:00-20:00", note: "" },
        { day: "Day 3", time: "外出不可", note: "昼食持参。忘れた人は運営スタッフが代わりに買いに行く" },
    ];

    outingData.forEach((item) => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: rightX + 0.04, y: ry, w: 0.65, h: 0.24,
            fill: { color: C.baseDark }, rectRadius: 0.04,
        });
        slide.addText(item.day, {
            x: rightX + 0.04, y: ry, w: 0.65, h: 0.24,
            fontSize: 9, fontFace: FONT_NUM, color: C.white, bold: true,
            align: "center", valign: "middle",
        });
        slide.addText(item.time, {
            x: rightX + 0.76, y: ry, w: colW - 0.80, h: 0.24,
            fontSize: 10, fontFace: FONT, color: C.textMain, bold: true,
            valign: "middle",
        });
        ry += 0.30;

        if (item.note) {
            slide.addText(`※ ${item.note}`, {
                x: rightX + 0.10, y: ry, w: colW - 0.14, h: 0.35,
                fontSize: 8, fontFace: FONT, color: C.accentRed,
            });
            ry += 0.38;
        }
    });

    // ===== 区切り線 =====
    const dividerY = Math.max(y, ry) + 0.08;
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: dividerY, w: CW, h: 0.015,
        fill: { color: C.gray200 },
    });

    // ===== 下部: 提出物一覧 =====
    let sy = dividerY + 0.18;
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: sy, w: 0.05, h: 0.25,
        fill: { color: C.accentRed },
    });
    slide.addText("📋 提出物一覧", {
        x: MX + 0.12, y: sy, w: CW - 0.12, h: 0.25,
        fontSize: 13, fontFace: FONT, color: C.baseDark, bold: true,
        valign: "middle",
    });
    sy += 0.35;

    const days = [
        {
            label: "Day 1",
            items: [
                "-13:30　コンサル社会人希望調査フォーム締切",
                "-15:30　コンサル状況共有フォーム締切",
                "-18:30　コンサル事後アンケート（全員提出）",
            ],
        },
        {
            label: "Day 2",
            items: [
                "-12:00　事業計画書提出フォーム",
                "-13:30　プレゼン資料提出フォーム",
                "-16:20　予選参加者審査フォーム（全員提出）",
                "-18:50　決勝参加者審査フォーム（全員提出）",
                "-21:15　総合事後アンケート（全員提出）",
            ],
        },
    ];

    // 提出物を2カラムで配置（Day1左、Day2右）
    const subColW = (CW - 0.15) / 2;
    days.forEach((day, di) => {
        const sx = MX + di * (subColW + 0.15);

        slide.addShape(pptx.ShapeType.roundRect, {
            x: sx, y: sy, w: 0.7, h: 0.24,
            fill: { color: C.baseDark }, rectRadius: 0.04,
        });
        slide.addText(day.label, {
            x: sx, y: sy, w: 0.7, h: 0.24,
            fontSize: 10, fontFace: FONT_NUM, color: C.white, bold: true,
            align: "center", valign: "middle",
        });

        let itemY = sy + 0.30;
        day.items.forEach((item) => {
            const match = item.match(/^(-[\d:]+)\s+(.+)$/);
            if (match) {
                slide.addText([
                    { text: `${match[1]}  `, options: { fontSize: 9, fontFace: FONT_NUM, color: C.baseDark, bold: true } },
                    { text: match[2], options: { fontSize: 9, fontFace: FONT, color: C.textMain } },
                ], {
                    x: sx + 0.06, y: itemY, w: subColW - 0.10, h: 0.22,
                    valign: "middle",
                });
            } else {
                slide.addText(`• ${item}`, {
                    x: sx + 0.06, y: itemY, w: subColW - 0.10, h: 0.22,
                    fontSize: 9, fontFace: FONT, color: C.textMain,
                    valign: "middle",
                });
            }
            itemY += 0.24;
        });
    });

    addFooter(pptx, slide);
}

// ============================================================
//  固定スライド7: ヒアリング実施時の注意事項
// ============================================================
function buildHearingRulesSlide(pptx: PptxGenJS) {
    const slide = pptx.addSlide();
    slide.background = { color: C.bgColor };

    // ヘッダー
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: W, h: HEADER_Y,
        fill: { color: C.accentRed },
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.35, w: 0.07, h: 0.4,
        fill: { color: C.accentRed },
    });
    slide.addText("⚠ ヒアリング実施時の注意事項", {
        x: MX + 0.18, y: 0.3, w: CW - 0.5, h: 0.5,
        fontSize: 18, fontFace: FONT, color: C.accentRed, bold: true,
        valign: "middle",
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.95, w: CW, h: 0.015,
        fill: { color: C.accentRed },
    });

    const rules = [
        {
            title: "正式な名乗りの統一と義務化",
            body: "ヒアリングを実施する際には必ず「東京大学の学生団体が主催するビジネスコンテストでのヒアリングとして」という団体の所属と目的を正確に表す文言を使用することを義務付ける。",
            theme: "danger" as const,
        },
        {
            title: "虚偽情報の使用の厳禁",
            body: "ヒアリングの実施にあたり、所属（大学、学部、学年）や目的等、いかなる情報においても誤った情報（虚偽の情報）を相手方に提示することを厳に禁止する。違反が発覚した場合は厳正に対処する。",
            theme: "danger" as const,
        },
        {
            title: "メールによるヒアリングの禁止",
            body: "２日間の土日開催であることを考慮し、本コンテストではメールによるヒアリングは実施しない。",
            theme: "warning" as const,
        },
        {
            title: "グラウンドルールへの配慮",
            body: "TSが場所を変えて実施するなど工夫すること。",
            theme: "info" as const,
        },
    ];

    let y = 1.2;

    rules.forEach((rule) => {
        const theme = ALERT_THEMES[rule.theme];
        const cardH = 1.4;

        // カード背景
        slide.addShape(pptx.ShapeType.roundRect, {
            x: MX, y, w: CW, h: cardH,
            fill: { color: theme.bg },
            rectRadius: 0.08,
        });

        // 左アクセントバー
        slide.addShape(pptx.ShapeType.rect, {
            x: MX, y, w: 0.06, h: cardH,
            fill: { color: theme.accent },
        });

        // タイトル
        slide.addText(`⚠ ${rule.title}`, {
            x: MX + 0.18, y: y + 0.08, w: CW - 0.30, h: 0.32,
            fontSize: 14, fontFace: FONT, color: theme.text, bold: true,
            valign: "middle",
        });

        // 本文
        slide.addText(rule.body, {
            x: MX + 0.18, y: y + 0.44, w: CW - 0.30, h: cardH - 0.56,
            fontSize: 11, fontFace: FONT, color: C.textMain,
            valign: "top",
        });

        y += cardH + 0.12;
    });

    addFooter(pptx, slide);
}

// ============================================================
//  共通ヘルパー: チャンネルカード描画（2カラムグリッド）
// ============================================================
function renderChannelGrid(
    pptx: PptxGenJS,
    slide: PptxGenJS.Slide,
    channels: { name: string; desc: string }[],
    startY: number
) {
    const colW = (CW - 0.15) / 2;
    const cardH = 0.65;
    const gap = 0.08;

    channels.forEach((ch, i) => {
        const col = i % 2;
        const row = Math.floor(i / 2);
        const x = MX + col * (colW + 0.15);
        const y = startY + row * (cardH + gap);

        // カード背景
        slide.addShape(pptx.ShapeType.roundRect, {
            x, y, w: colW, h: cardH,
            fill: { color: C.bgTableHead },
            rectRadius: 0.06,
        });

        // 左アクセントライン
        slide.addShape(pptx.ShapeType.rect, {
            x, y, w: 0.04, h: cardH,
            fill: { color: C.baseBlue },
        });

        // チャンネル名
        slide.addText(ch.name, {
            x: x + 0.1, y, w: colW - 0.14, h: 0.22,
            fontSize: 9, fontFace: FONT, color: C.baseBlue, bold: true,
            valign: "bottom",
        });

        // 説明
        slide.addText(ch.desc, {
            x: x + 0.1, y: y + 0.24, w: colW - 0.14, h: cardH - 0.28,
            fontSize: 8, fontFace: FONT, color: C.textMain,
            valign: "top",
        });
    });
}

// ============================================================
//  共通ヘルパー: チャンネルスライドヘッダー
// ============================================================
function addChannelSlideHeader(pptx: PptxGenJS, slide: PptxGenJS.Slide, title: string) {
    slide.background = { color: C.bgColor };
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: W, h: HEADER_Y,
        fill: { color: C.baseDark },
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.30, w: 0.07, h: 0.45,
        fill: { color: C.baseBlue },
    });
    slide.addText(title, {
        x: MX + 0.18, y: 0.30, w: CW - 0.5, h: 0.45,
        fontSize: 16, fontFace: FONT, color: C.textMain, bold: true,
        valign: "middle",
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.90, w: CW, h: 0.015,
        fill: { color: C.gray200 },
    });
}

// ============================================================
//  固定スライド8: パブリックチャンネル
// ============================================================
function buildSlackPublicSlide(pptx: PptxGenJS) {
    const slide = pptx.addSlide();
    addChannelSlideHeader(pptx, slide, "パブリックチャンネル（参加者・KINGer全員）");

    const channels = [
        { name: "#001_general_全体連絡", desc: "参加者・KINGer全体に対する連絡" },
        { name: "#002_general_質問部屋", desc: "参加者からの質問受付（※KINGerは #103 を利用）" },
        { name: "#003_general_資料提出", desc: "各種提出するもの、フォーム等の連絡を行う" },
        { name: "#004_general_遺失物連絡", desc: "運営スタッフから遺失物等の連絡を行う" },
        { name: "#005_general_ヒアリング", desc: "他参加者へのヒアリング調査、アンケートフォームの展開などに利用可能" },
        { name: "#006_general_告知", desc: "協賛企業のイベント告知などを行う" },
        { name: "#007_general_質疑応答", desc: "予選審査の際、参加者・TSはフォーマットに従って質疑応答可能" },
    ];

    renderChannelGrid(pptx, slide, channels, 1.1);
    addFooter(pptx, slide);
}

// ============================================================
//  固定スライド9: プライベート（100番台）
// ============================================================
function buildSlackPrivate100Slide(pptx: PptxGenJS) {
    const slide = pptx.addSlide();
    addChannelSlideHeader(pptx, slide, "プライベートチャンネル（KINGerのみ）");

    const channels = [
        { name: "#101_staff_全体連絡", desc: "KINGer全員に対する連絡" },
        { name: "#102_staff_緊急連絡", desc: "スタッフ向けの緊急性の高い連絡を行う" },
        { name: "#103_staff_質問部屋", desc: "KINGerからの質問に対応する" },
        { name: "#104_staff_体調管理", desc: "スタッフ及び参加者の体調連絡" },
        { name: "#105_staff_欠席早退遅刻連絡", desc: "スタッフ及び参加者の出席早退遅刻連絡" },
        { name: "#106_staff_入退場連絡", desc: "スタッフおよび参加者の入退場連絡" },
        { name: "#107_staff_巡回呼び出し", desc: "巡回TSを参加者とTSが呼ぶ部屋" },
        { name: "#108_staff_ts", desc: "TS向けの連絡（プランニングに関する連絡）" },
        { name: "#109_staff_巡回スタッフ", desc: "巡回同士で話す部屋" },
        { name: "#110_staff_写真投稿", desc: "写真を格納する" },
    ];

    renderChannelGrid(pptx, slide, channels, 1.1);
    addFooter(pptx, slide);
}

// ============================================================
//  固定スライド10: プライベート（200〜300番台）
// ============================================================
function buildSlackPrivate200Slide(pptx: PptxGenJS) {
    const slide = pptx.addSlide();
    addChannelSlideHeader(pptx, slide, "プライベートチャンネル（200〜300番台）");

    const channels = [
        { name: "#201_staff_司会レス", desc: "司会・レス・渉外局との連絡" },
        { name: "#202_staff_審査", desc: "審査（予選・決勝）における連絡" },
        { name: "#203_staff_運営スタッフ", desc: "運営スタッフの連絡を行う" },
        { name: "#204_staff_運営チーム連携", desc: "運営スタッフの連絡" },
        { name: "#205_staff_カメラマン", desc: "カメラ関係の連絡" },
        { name: "#206_staff_印刷", desc: "印刷関連の連絡" },
        { name: "#301_staff_社会人対応", desc: "社会人対応関連の連絡" },
        { name: "#302_staff_物品協賛", desc: "物品協賛物連絡" },
    ];

    renderChannelGrid(pptx, slide, channels, 1.1);
    addFooter(pptx, slide);
}

// ============================================================
//  固定スライド11: プライベート（400〜500番台）
// ============================================================
function buildSlackPrivate400Slide(pptx: PptxGenJS) {
    const slide = pptx.addSlide();
    addChannelSlideHeader(pptx, slide, "プライベートチャンネル（400〜500番台）");

    const channels = [
        { name: "#401_staff_代表直通", desc: "代表への連絡" },
        { name: "#402_staff_幹部直通", desc: "幹部への連絡" },
        { name: "#403_staff_渉外局幹部直通", desc: "渉外局幹部・社会人に関しての連絡" },
        { name: "#404_staff_役員直通", desc: "財務・備品・印刷・法務に関して" },
        { name: "#405_staff_運営統括局幹部直通", desc: "運営全体に関して" },
        { name: "#406_staff_開発局幹部直通", desc: "ケース・審査基準・巡回・TSに関して" },
        { name: "#407_staff_れあ直通", desc: "参加者対応に関して" },
        { name: "#501_staff_反省_全体", desc: "反省点を書き留める" },
        { name: "#502_staff_ぼやき", desc: "みんなでぼやこう🫠" },
        { name: "#503_staff_ts_ぼやき", desc: "ts関係でぼやこう🫠" },
    ];

    renderChannelGrid(pptx, slide, channels, 1.1);
    addFooter(pptx, slide);
}

// ============================================================
//  固定スライド12: 体調管理の使い方
// ============================================================
function buildSlackHealthSlide(pptx: PptxGenJS) {
    const slide = pptx.addSlide();
    addChannelSlideHeader(pptx, slide, "具体的な使い方：#104_staff_体調管理");

    const cardY = 1.15;
    const cardH = 5.0;

    // メインカード（BgCardInfo）
    slide.addShape(pptx.ShapeType.roundRect, {
        x: MX, y: cardY, w: CW, h: cardH,
        fill: { color: C.bgCardInfo },
        rectRadius: 0.1,
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: cardY, w: 0.06, h: cardH,
        fill: { color: C.baseBlue },
    });

    let y = cardY + 0.15;

    // 用途
    slide.addText([
        { text: "用途：", options: { fontSize: 12, fontFace: FONT, color: C.baseDark, bold: true } },
        { text: "参加者・TSの体調管理を行うチャンネル", options: { fontSize: 12, fontFace: FONT, color: C.textMain } },
    ], {
        x: MX + 0.18, y, w: CW - 0.30, h: 0.30,
        valign: "middle",
    });
    y += 0.40;

    // ルール
    slide.addText([
        { text: "ルール：", options: { fontSize: 12, fontFace: FONT, color: C.baseDark, bold: true } },
        { text: "ワークフローに従い報告すること", options: { fontSize: 12, fontFace: FONT, color: C.textMain } },
    ], {
        x: MX + 0.18, y, w: CW - 0.30, h: 0.30,
        valign: "middle",
    });
    y += 0.50;

    // 入力項目見出し
    slide.addShape(pptx.ShapeType.rect, {
        x: MX + 0.14, y, w: 0.05, h: 0.25,
        fill: { color: C.baseDark },
    });
    slide.addText("📋 ワークフロー入力項目", {
        x: MX + 0.26, y, w: CW - 0.40, h: 0.25,
        fontSize: 13, fontFace: FONT, color: C.baseDark, bold: true,
        valign: "middle",
    });
    y += 0.40;

    // 入力項目
    const inputFields = [
        { label: "ワークフロー入力者", icon: "✏️" },
        { label: "チーム名", icon: "👥" },
        { label: "体調不良対象者の名前", icon: "🏷️" },
        { label: "役職", icon: "📌" },
        { label: "参加者", icon: "👤" },
        { label: "望んでいる対応", icon: "🩹" },
        { label: "症状", icon: "🤒" },
    ];

    inputFields.forEach((field) => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: MX + 0.18, y, w: CW - 0.36, h: 0.40,
            fill: { color: C.white },
            rectRadius: 0.05,
        });
        slide.addText(`${field.icon}  ${field.label}`, {
            x: MX + 0.30, y, w: CW - 0.60, h: 0.40,
            fontSize: 11, fontFace: FONT, color: C.textMain, bold: true,
            valign: "middle",
        });
        y += 0.46;
    });

    addFooter(pptx, slide);
}

// ============================================================
//  固定スライド13: 欠席早退遅刻連絡の使い方
// ============================================================
function buildSlackAttendanceSlide(pptx: PptxGenJS) {
    const slide = pptx.addSlide();
    addChannelSlideHeader(pptx, slide, "具体的な使い方：#105_staff_欠席早退遅刻連絡");

    const colW = (CW - 0.12) / 2;
    const blockH = 2.4;
    const gap = 0.12;
    const startY = 1.10;

    const blocks = [
        {
            title: "欠席の場合",
            bg: C.bgCardInfo,
            accent: C.baseBlue,
            items: [
                "参加者：TSに連絡 → TSがワークフローで報告",
                "KINGer：本人がワークフローで報告",
            ],
        },
        {
            title: "遅刻の場合",
            bg: C.bgCardWarning,
            accent: C.amber,
            items: [
                "参加者：TSに連絡 → TSが到着予想時刻を確認し「到着次第運営スタッフが迎えに行く」旨を伝える → ワークフローで報告 → 運営スタッフが迎え",
                "KINGer：本人がワークフローで報告 → 到着後ワークフローにスタンプ",
            ],
        },
        {
            title: "早退の場合",
            bg: C.bgCardInfo,
            accent: C.baseBlue,
            items: [
                "参加者：TSに連絡 → TSがワークフローで報告",
                "KINGer：本人がワークフローで報告",
            ],
        },
        {
            title: "⚠️ 無断遅刻・無断欠席の場合",
            bg: C.bgCardDanger,
            accent: C.accentRed,
            items: [
                "TSが #102_staff_緊急連絡 にて @運営統括局幹部 @久保井 にメンション → 未着のTS・参加者の氏名を報告",
                "その後、役員・運営スタッフ・久保井が直接連絡する",
            ],
        },
    ];

    blocks.forEach((block, i) => {
        const col = i % 2;
        const row = Math.floor(i / 2);
        const x = MX + col * (colW + gap);
        const y = startY + row * (blockH + gap);

        // カード背景
        slide.addShape(pptx.ShapeType.roundRect, {
            x, y, w: colW, h: blockH,
            fill: { color: block.bg },
            rectRadius: 0.08,
        });

        // 左アクセントバー
        slide.addShape(pptx.ShapeType.rect, {
            x, y, w: 0.05, h: blockH,
            fill: { color: block.accent },
        });

        // タイトル
        slide.addText(block.title, {
            x: x + 0.12, y: y + 0.08, w: colW - 0.20, h: 0.30,
            fontSize: 12, fontFace: FONT, color: block.accent, bold: true,
            valign: "middle",
        });

        // 区切り線
        slide.addShape(pptx.ShapeType.rect, {
            x: x + 0.10, y: y + 0.42, w: colW - 0.20, h: 0.01,
            fill: { color: block.accent },
        });

        // 項目
        let itemY = y + 0.52;
        block.items.forEach((item) => {
            const lineCount = Math.ceil(item.length / 18);
            const itemH = lineCount * 0.18 + 0.15;
            slide.addText(`▸ ${item}`, {
                x: x + 0.12, y: itemY, w: colW - 0.24, h: itemH,
                fontSize: 9, fontFace: FONT, color: C.textMain,
                valign: "top",
            });
            itemY += itemH + 0.06;
        });
    });

    addFooter(pptx, slide);
}
// ============================================================
//  統合スライド: パブリック（上部）+ プライベート100番台（下部）
// ============================================================
function buildSlackCombinedPublicAnd100Slide(pptx: PptxGenJS) {
    const slide = pptx.addSlide();
    slide.background = { color: C.bgColor };

    // ヘッダー
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: W, h: HEADER_Y,
        fill: { color: C.baseDark },
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.25, w: 0.07, h: 0.45,
        fill: { color: C.baseBlue },
    });
    slide.addText("Slackチャンネル：パブリック & プライベート（100番台）", {
        x: MX + 0.18, y: 0.22, w: CW - 0.5, h: 0.50,
        fontSize: 14, fontFace: FONT, color: C.textMain, bold: true,
        valign: "middle",
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.85, w: CW, h: 0.015,
        fill: { color: C.gray200 },
    });

    // 上部: パブリック
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 1.0, w: 0.04, h: 0.20,
        fill: { color: C.green },
    });
    slide.addText("パブリックチャンネル（参加者・KINGer全員）", {
        x: MX + 0.10, y: 1.0, w: CW - 0.10, h: 0.20,
        fontSize: 10, fontFace: FONT, color: C.baseDark, bold: true,
        valign: "middle",
    });

    const pubChannels = [
        { name: "#001_general_全体連絡", desc: "参加者・KINGer全体に対する連絡" },
        { name: "#002_general_質問部屋", desc: "参加者からの質問受付" },
        { name: "#003_general_資料提出", desc: "フォーム等の連絡を行う" },
        { name: "#004_general_遺失物連絡", desc: "遺失物等の連絡" },
        { name: "#005_general_ヒアリング", desc: "ヒアリング・アンケート" },
        { name: "#006_general_告知", desc: "協賛企業のイベント告知" },
        { name: "#007_general_質疑応答", desc: "予選の質疑応答" },
    ];
    renderChannelGrid(pptx, slide, pubChannels, 1.28);

    // 区切り線
    const midY = 1.28 + Math.ceil(pubChannels.length / 2) * 0.73 + 0.08;
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: midY, w: CW, h: 0.015,
        fill: { color: C.gray200 },
    });

    // 下部: プライベート100番台
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: midY + 0.10, w: 0.04, h: 0.20,
        fill: { color: C.baseDark },
    });
    slide.addText("プライベートチャンネル（KINGerのみ・100番台）", {
        x: MX + 0.10, y: midY + 0.10, w: CW - 0.10, h: 0.20,
        fontSize: 10, fontFace: FONT, color: C.baseDark, bold: true,
        valign: "middle",
    });

    const priv100 = [
        { name: "#101_staff_全体連絡", desc: "KINGer全員に対する連絡" },
        { name: "#102_staff_緊急連絡", desc: "緊急性の高い連絡" },
        { name: "#103_staff_質問部屋", desc: "KINGerの質問対応" },
        { name: "#104_staff_体調管理", desc: "体調連絡" },
        { name: "#105_staff_欠席早退遅刻", desc: "出席関連の連絡" },
        { name: "#106_staff_入退場連絡", desc: "入退場連絡" },
        { name: "#107_staff_巡回呼び出し", desc: "巡回TSの呼び出し" },
        { name: "#108_staff_ts", desc: "TS向け連絡" },
        { name: "#109_staff_巡回スタッフ", desc: "巡回同士の連絡" },
        { name: "#110_staff_写真投稿", desc: "写真格納" },
    ];
    renderChannelGrid(pptx, slide, priv100, midY + 0.38);

    addFooter(pptx, slide);
}

// ============================================================
//  統合スライド: プライベート200〜300番台（上部）+ 400〜500番台（下部）
// ============================================================
function buildSlackCombined200to500Slide(pptx: PptxGenJS) {
    const slide = pptx.addSlide();
    slide.background = { color: C.bgColor };

    // ヘッダー
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: W, h: HEADER_Y,
        fill: { color: C.baseDark },
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.25, w: 0.07, h: 0.45,
        fill: { color: C.baseBlue },
    });
    slide.addText("Slackチャンネル：プライベート（200〜500番台）", {
        x: MX + 0.18, y: 0.22, w: CW - 0.5, h: 0.50,
        fontSize: 14, fontFace: FONT, color: C.textMain, bold: true,
        valign: "middle",
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.85, w: CW, h: 0.015,
        fill: { color: C.gray200 },
    });

    // 上部: 200〜300番台
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 1.0, w: 0.04, h: 0.20,
        fill: { color: C.purple },
    });
    slide.addText("200〜300番台", {
        x: MX + 0.10, y: 1.0, w: CW - 0.10, h: 0.20,
        fontSize: 10, fontFace: FONT, color: C.baseDark, bold: true,
        valign: "middle",
    });

    const priv200 = [
        { name: "#201_staff_司会レス", desc: "司会・レス・渉外局" },
        { name: "#202_staff_審査", desc: "予選・決勝の連絡" },
        { name: "#203_staff_運営スタッフ", desc: "運営スタッフ連絡" },
        { name: "#204_staff_運営チーム連携", desc: "運営チーム連絡" },
        { name: "#205_staff_カメラマン", desc: "カメラ関係" },
        { name: "#206_staff_印刷", desc: "印刷関連" },
        { name: "#301_staff_社会人対応", desc: "社会人対応" },
        { name: "#302_staff_物品協賛", desc: "物品協賛" },
    ];
    renderChannelGrid(pptx, slide, priv200, 1.28);

    // 区切り線
    const midY = 1.28 + Math.ceil(priv200.length / 2) * 0.73 + 0.08;
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: midY, w: CW, h: 0.015,
        fill: { color: C.gray200 },
    });

    // 下部: 400〜500番台
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: midY + 0.10, w: 0.04, h: 0.20,
        fill: { color: C.teal },
    });
    slide.addText("400〜500番台", {
        x: MX + 0.10, y: midY + 0.10, w: CW - 0.10, h: 0.20,
        fontSize: 10, fontFace: FONT, color: C.baseDark, bold: true,
        valign: "middle",
    });

    const priv400 = [
        { name: "#401_staff_代表直通", desc: "代表への連絡" },
        { name: "#402_staff_幹部直通", desc: "幹部への連絡" },
        { name: "#403_staff_渉外局幹部直通", desc: "渉外局幹部・社会人" },
        { name: "#404_staff_役員直通", desc: "財務・備品・法務" },
        { name: "#405_staff_運営統括局幹部", desc: "運営全体" },
        { name: "#406_staff_開発局幹部", desc: "ケース・審査基準" },
        { name: "#407_staff_れあ直通", desc: "参加者対応" },
        { name: "#501_staff_反省_全体", desc: "反省点の記録" },
        { name: "#502_staff_ぼやき", desc: "みんなでぼやこう🫠" },
        { name: "#503_staff_ts_ぼやき", desc: "ts関係のぼやき🫠" },
    ];
    renderChannelGrid(pptx, slide, priv400, midY + 0.38);

    addFooter(pptx, slide);
}

// ============================================================
//  固定スライド7: Slackチャンネル説明（全情報を1枚に圧縮）
//  3カラムグリッド + 極小フォント + セクション分けで全35chを配置
// ============================================================
function buildSlackChannelsCompactSlide(pptx: PptxGenJS) {
    const slide = pptx.addSlide();
    slide.background = { color: C.bgColor };

    // ヘッダー帯
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: W, h: 0.75,
        fill: { color: C.baseDark },
    });
    slide.addText("Slackチャンネルガイド", {
        x: 0, y: 0, w: W, h: 0.75,
        fontSize: 22, fontFace: FONT, color: C.white, bold: true,
        align: "center", valign: "middle",
    });

    const col3W = (CW - 0.16) / 3;
    const colXs = [MX, MX + col3W + 0.08, MX + (col3W + 0.08) * 2];
    const lineH = 0.17;
    const sectionGap = 0.08;

    // チャンネル描画ヘルパー
    const drawChannel = (x: number, y: number, name: string, desc: string) => {
        slide.addText(name, {
            x: x + 0.04, y, w: col3W - 0.08, h: lineH * 0.5,
            fontSize: 6.5, fontFace: FONT, color: C.baseBlue, bold: true,
            valign: "bottom",
        });
        slide.addText(desc, {
            x: x + 0.04, y: y + lineH * 0.48, w: col3W - 0.08, h: lineH * 0.52,
            fontSize: 5.5, fontFace: FONT, color: C.textSub,
            valign: "top",
        });
    };

    // セクション描画ヘルパー
    const drawSection = (colIdx: number, startY: number, title: string, channels: { n: string; d: string }[]): number => {
        const x = colXs[colIdx];
        // セクションタイトル
        slide.addShape(pptx.ShapeType.rect, {
            x, y: startY, w: 0.03, h: 0.15,
            fill: { color: C.baseDark },
        });
        slide.addText(title, {
            x: x + 0.06, y: startY, w: col3W - 0.06, h: 0.15,
            fontSize: 7, fontFace: FONT, color: C.baseDark, bold: true,
            valign: "middle",
        });

        let cy = startY + 0.18;
        channels.forEach((ch) => {
            drawChannel(x, cy, ch.n, ch.d);
            cy += lineH;
        });
        return cy + sectionGap;
    };

    // ===== カラム1: パブリック =====
    const pubChannels = [
        { n: "#001_general_全体連絡", d: "参加者・KINGer全体の連絡" },
        { n: "#002_general_質問部屋", d: "参加者からの質問受付" },
        { n: "#003_general_資料提出", d: "フォーム等の連絡" },
        { n: "#004_general_遺失物連絡", d: "遺失物の連絡" },
        { n: "#005_general_ヒアリング", d: "ヒアリング・アンケート" },
        { n: "#006_general_告知", d: "協賛企業のイベント告知" },
        { n: "#007_general_質疑応答", d: "予選時の質疑応答" },
    ];
    let nextY1 = drawSection(0, 0.90, "パブリック（全員）", pubChannels);

    // プライベート100番台をカラム1の続きに
    const priv100 = [
        { n: "#101_staff_全体連絡", d: "KINGer全員の連絡" },
        { n: "#102_staff_緊急連絡", d: "緊急性の高い連絡" },
        { n: "#103_staff_質問部屋", d: "KINGerの質問対応" },
        { n: "#104_staff_体調管理", d: "体調連絡" },
        { n: "#105_staff_欠席早退遅刻", d: "出席関連の連絡" },
    ];
    drawSection(0, nextY1, "プライベート 100番台", priv100);

    // ===== カラム2: プライベート100番台（続き）+ 200〜300番台 =====
    const priv100b = [
        { n: "#106_staff_入退場連絡", d: "入退場連絡" },
        { n: "#107_staff_巡回呼び出し", d: "巡回TSの呼び出し" },
        { n: "#108_staff_ts", d: "TS向け連絡" },
        { n: "#109_staff_巡回スタッフ", d: "巡回同士の連絡" },
        { n: "#110_staff_写真投稿", d: "写真格納" },
    ];
    let nextY2 = drawSection(1, 0.90, "プライベート 100番台（続）", priv100b);

    const priv200 = [
        { n: "#201_staff_司会レス", d: "司会・レス・渉外局" },
        { n: "#202_staff_審査", d: "予選・決勝の連絡" },
        { n: "#203_staff_運営スタッフ", d: "運営スタッフ連絡" },
        { n: "#204_staff_運営チーム連携", d: "運営チーム連絡" },
        { n: "#205_staff_カメラマン", d: "カメラ関係" },
        { n: "#206_staff_印刷", d: "印刷関連" },
        { n: "#301_staff_社会人対応", d: "社会人対応" },
        { n: "#302_staff_物品協賛", d: "物品協賛" },
    ];
    drawSection(1, nextY2, "200〜300番台", priv200);

    // ===== カラム3: 400〜500番台 =====
    const priv400 = [
        { n: "#401_staff_代表直通", d: "代表への連絡" },
        { n: "#402_staff_幹部直通", d: "幹部への連絡" },
        { n: "#403_staff_渉外局幹部", d: "渉外局幹部・社会人" },
        { n: "#404_staff_役員直通", d: "財務・備品・法務" },
        { n: "#405_staff_運営統括局幹部", d: "運営全体" },
        { n: "#406_staff_開発局幹部", d: "ケース・審査基準" },
        { n: "#407_staff_れあ直通", d: "参加者対応" },
        { n: "#501_staff_反省_全体", d: "反省点の記録" },
        { n: "#502_staff_ぼやき", d: "みんなでぼやこう🫠" },
        { n: "#503_staff_ts_ぼやき", d: "ts関係のぼやき🫠" },
    ];
    drawSection(2, 0.90, "400〜500番台", priv400);

    addFooter(pptx, slide);
}

// ============================================================
//  スライドヘッダー（2行レイアウト: コンテスト名 + コンテンツ名）
// ============================================================
function addSlideHeader(
    pptx: PptxGenJS,
    slide: PptxGenJS.Slide,
    info: BasicInfo,
    _type: "timeline" | "notes",
    _num: number,
    _total: number
) {
    // 上バー
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: W, h: HEADER_Y,
        fill: { color: C.baseDark },
    });

    // 縦バー
    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 0.25, w: 0.07, h: 0.45,
        fill: { color: C.baseDark },
    });

    // コンテンツ名のみ（20pt・太字）
    slide.addText(info.content_name || "名称未設定", {
        x: MX + 0.18, y: 0.22, w: CW - 0.5, h: 0.50,
        fontSize: 20, fontFace: FONT, color: C.textMain, bold: true,
        valign: "middle",
    });

    // 日付・時間・場所
    const infoItems = [
        info.date ? `📅 ${info.date}` : "",
        info.time ? `🕐 ${info.time}` : "",
        info.location ? `📍 ${info.location}` : "",
    ].filter(s => s.length > 0);

    if (infoItems.length > 0) {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: MX, y: 0.88, w: CW, h: 0.28,
            fill: { color: C.bgTableHead },
            line: { color: C.gray200, width: 0.3 },
            rectRadius: 0.05,
        });
        infoItems.forEach((item, i) => {
            slide.addText(item, {
                x: MX + 0.12 + i * 2.05, y: 0.88, w: 2.0, h: 0.28,
                fontSize: 8.5, fontFace: FONT, color: C.textSub, valign: "middle",
            });
        });
    }

    slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: 1.3, w: CW, h: 0.015,
        fill: { color: C.gray200 },
    });
}

// ============================================================
//  フッター
// ============================================================
function addFooter(pptx: PptxGenJS, slide: PptxGenJS.Slide) {
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: H - FOOTER_H, w: 0.06, h: FOOTER_H,
        fill: { color: C.baseBlue },
    });
    slide.addText("Event Operations Manual", {
        x: 0.2, y: H - FOOTER_H, w: 3, h: FOOTER_H,
        fontSize: 7, fontFace: FONT_NUM, color: C.gray400, valign: "middle",
    });
}

// ============================================================
//  続きスライドの生成（スライドカウンタ付き）
// ============================================================
function addContinuationSlide(
    pptx: PptxGenJS,
    info: BasicInfo,
    type: "timeline" | "notes",
    num: number,
    total: number,
    ctx: SlideContext
): SlideContext {
    // 3枚制限チェック
    if (ctx.slideCount >= MAX_SLIDES_PER_CONTENT) {
        // 制限に達している場合は現在のスライドのまま続行
        return ctx;
    }

    const slide = pptx.addSlide();
    slide.background = { color: C.bgColor };
    addSlideHeader(pptx, slide, info, type, num, total);
    addFooter(pptx, slide);

    return {
        slide,
        y: CONTINUATION_TOP,
        slideCount: ctx.slideCount + 1,
    };
}

// ============================================================
//  メイン描画振り分け（3枚制限付き）
// ============================================================
function buildContentSlides(
    pptx: PptxGenJS,
    content: ContentData,
    num: number,
    total: number,
    layoutImages?: string[]
) {
    const rawInfo = content.basic_info;
    // content_nameから年号以降を切り出す（例: "Spring Business Contest 2026 Day0" → "Day0"）
    const trimmedName = rawInfo.content_name
        ? rawInfo.content_name.replace(/^.*\d{4}\s*[\/\-]?\s*/i, "").trim() || rawInfo.content_name
        : rawInfo.content_name;
    const info: BasicInfo = { ...rawInfo, content_name: trimmedName };
    const hasTimeline = content.timeline && content.timeline.length > 0;
    const hasNotes = content.important_notes && content.important_notes.length > 0;
    const hasAlerts = content.critical_timeline_alerts && content.critical_timeline_alerts.length > 0;

    // レイアウト画像の取得
    let layoutBase64: string | null = null;
    if (
        content.layout_page_number != null &&
        layoutImages &&
        layoutImages.length > 0 &&
        content.layout_page_number >= 1 &&
        content.layout_page_number <= layoutImages.length
    ) {
        layoutBase64 = layoutImages[content.layout_page_number - 1];
    }

    // 最初のスライドを生成
    const slide = pptx.addSlide();
    slide.background = { color: C.bgColor };
    addSlideHeader(pptx, slide, info, "timeline", num, total);
    addFooter(pptx, slide);

    let ctx: SlideContext = {
        slide,
        y: CONTENT_TOP,
        slideCount: 1,
    };

    // 1) タイムライン描画
    if (hasTimeline) {
        ctx = renderTimeline(pptx, ctx, content.timeline, info, num, total);
    }

    // 2) 重要事項カード描画
    if (hasNotes) {
        ctx.y += 0.12;
        if (ctx.y + 0.6 > PAGE_BOTTOM && ctx.slideCount < MAX_SLIDES_PER_CONTENT) {
            ctx = addContinuationSlide(pptx, info, "notes", num, total, ctx);
        }
        ctx = renderImportantNotes(pptx, ctx, content.important_notes, info, num, total);
    }

    // 3) クリティカルアラート描画
    if (hasAlerts) {
        ctx.y += 0.1;
        if (ctx.y + 0.6 > PAGE_BOTTOM && ctx.slideCount < MAX_SLIDES_PER_CONTENT) {
            ctx = addContinuationSlide(pptx, info, "notes", num, total, ctx);
        }
        ctx = renderCriticalAlerts(pptx, ctx, content.critical_timeline_alerts, info, num, total);
    }

    // 4) レイアウト画像の動的配置（専用スライドではなく空きスペースに配置）
    if (layoutBase64) {
        placeLayoutImage(pptx, ctx, layoutBase64, info, num, total);
    }
}

// ============================================================
//  タイムライン描画（3枚制限対応）
// ============================================================
function renderTimeline(
    pptx: PptxGenJS,
    ctx: SlideContext,
    timeline: ContentData["timeline"],
    info: BasicInfo,
    num: number,
    total: number,
): SlideContext {
    if (!timeline || timeline.length === 0) return ctx;

    const sanitized = timeline.filter(t => t && t.action.trim().length > 0);
    if (sanitized.length === 0) return ctx;

    const rowH = 0.32;
    const headerH = rowH;

    const col0W = CW * 0.14;
    const col1W = CW * 0.14;
    const col2W = CW * 0.72;

    const isImportantAction = (action: string): string | null => {
        const a = action.toLowerCase();
        if (a.includes("集合") || a.includes("準備") || a.includes("受付") || a.includes("搬入") || a.includes("リハ")) return C.amberSoft;
        if (a.includes("開始") || a.includes("開演") || a.includes("本番") || a.includes("開会")) return C.bgCardInfo;
        if (a.includes("終了") || a.includes("解散") || a.includes("撤収") || a.includes("閉会")) return C.bgCardDanger;
        return null;
    };

    const isImportantBorder = (action: string): string | null => {
        const a = action.toLowerCase();
        if (a.includes("集合") || a.includes("準備") || a.includes("受付") || a.includes("搬入")) return C.amber;
        if (a.includes("開始") || a.includes("開演") || a.includes("本番") || a.includes("開会")) return C.baseBlue;
        if (a.includes("終了") || a.includes("解散") || a.includes("撤収") || a.includes("閉会")) return C.accentRed;
        return null;
    };

    // モダンテーブルヘッダー（BaseDark青背景 + 白文字で目立たせる）
    const makeHeaderCell = (text: string): PptxGenJS.TableCell => ({
        text,
        options: {
            fontSize: 10,
            fontFace: FONT, color: C.white, bold: true,
            fill: { color: C.baseDark },
            valign: "middle", align: "center" as const,
            margin: [2, 4, 2, 4],
            border: [
                { type: "none", pt: 0, color: C.white },
                { type: "none", pt: 0, color: C.white },
                { type: "none", pt: 0, color: C.white },
                { type: "none", pt: 0, color: C.white },
            ],
        },
    });

    const makeDataCell = (text: string, tc: string, bold: boolean, align: "left" | "center" = "left", leftBorder?: string, isNumeric: boolean = false): PptxGenJS.TableCell => ({
        text,
        options: {
            fontSize: MIN_BODY_FONT,
            fontFace: isNumeric ? FONT_NUM : FONT, color: tc, bold: bold || isNumeric,
            fill: { color: C.white },
            valign: "middle", align,
            margin: align === "center" ? [2, 4, 2, 4] : [2, 8, 2, 8],
            border: [
                { type: "none", pt: 0, color: C.white },
                { type: "solid", pt: 0.5, color: C.borderLight },
                { type: "solid", pt: leftBorder ? 2.5 : 0, color: leftBorder || C.white },
                { type: "none", pt: 0, color: C.white },
            ],
        },
    });

    // title-bar風見出し（左縦線 + ラベル）
    ctx.slide.addShape(pptx.ShapeType.rect, {
        x: MX, y: ctx.y - 0.38, w: 0.05, h: 0.28,
        fill: { color: C.baseDark },
    });
    ctx.slide.addText("タイムライン", {
        x: MX + 0.12, y: ctx.y - 0.38, w: 2, h: 0.28,
        fontSize: 11, fontFace: FONT, color: C.textMain, bold: true,
        valign: "middle",
    });

    const tableRows: PptxGenJS.TableRow[] = [
        [
            makeHeaderCell("開始"),
            makeHeaderCell("終了"),
            { text: "アクション", options: { ...makeHeaderCell("アクション").options, align: "left" as const, margin: [2, 8, 2, 8] } },
        ],
        ...sanitized.map((item, idx) => {
            const highlight = isImportantAction(item.action);
            const borderColor = isImportantBorder(item.action) || undefined;
            const bold = !!highlight || idx === 0;
            // 1行目は青色背景で目立たせる
            const isFirstRow = idx === 0;
            const rowBg = isFirstRow ? C.bgCardInfo : C.white;
            const rowTextColor = isFirstRow ? C.baseDark : C.textMain;
            return [
                { text: item.start_time, options: { ...makeDataCell(item.start_time, rowTextColor, bold, "center", borderColor || (isFirstRow ? C.baseBlue : undefined), true).options, fill: { color: rowBg } } },
                { text: item.end_time, options: { ...makeDataCell(item.end_time, isFirstRow ? C.baseDark : C.textSub, isFirstRow, "center", undefined, true).options, fill: { color: rowBg } } },
                { text: item.action, options: { ...makeDataCell(item.action, rowTextColor, bold).options, fill: { color: rowBg } } },
            ];
        }),
    ];

    const totalTableH = headerH + rowH * sanitized.length;
    const availableH = PAGE_BOTTOM - ctx.y;

    // テーブルが1ページに収まる場合
    if (totalTableH <= availableH) {
        ctx.slide.addShape(pptx.ShapeType.roundRect, {
            x: MX - 0.04, y: ctx.y - 0.04, w: CW + 0.08, h: totalTableH + 0.08,
            fill: { color: C.white },
            shadow: { type: "outer", blur: 4, offset: 1, color: "000000", opacity: 0.06 },
            line: { color: C.gray200, width: 0.5 }, rectRadius: 0.08,
        });

        ctx.slide.addTable(tableRows, {
            x: MX, y: ctx.y, w: CW,
            colW: [col0W, col1W, col2W],
            rowH: [headerH, ...Array(sanitized.length).fill(rowH)],
            border: { type: "none", pt: 0, color: C.white },
        });

        // 凡例
        const legendY = ctx.y + totalTableH + 0.08;
        [
            { label: "● 集合・準備", color: C.amber },
            { label: "● 開始・本番", color: C.baseBlue },
            { label: "● 終了・解散", color: C.accentRed },
        ].forEach((item, i) => {
            ctx.slide.addText(item.label, {
                x: MX + i * 1.5, y: legendY, w: 1.4, h: 0.2,
                fontSize: 6.5, fontFace: FONT, color: item.color, bold: true,
            });
        });

        return { ...ctx, y: ctx.y + totalTableH + 0.32 };
    }

    // テーブルが長い場合: autoPageで自動改ページ（ただし3枚制限を意識）
    ctx.slide.addTable(tableRows, {
        x: MX, y: ctx.y, w: CW,
        colW: [col0W, col1W, col2W],
        rowH: [headerH, ...Array(sanitized.length).fill(rowH)],
        border: { type: "none", pt: 0, color: C.white },
        autoPage: true,
        autoPageRepeatHeader: true,
        autoPageLineWeight: 0.32,
        newSlideStartY: CONTINUATION_TOP,
    });

    // autoPageで複数ページに分割された推定
    const rowsPerPage = Math.floor(availableH / rowH);
    const totalPages = Math.ceil(sanitized.length / rowsPerPage);
    const newSlideCount = ctx.slideCount + Math.max(0, totalPages - 1);

    if (totalPages > 1) {
        const lastPageRows = sanitized.length - rowsPerPage * (totalPages - 1);
        return {
            slide: ctx.slide, // autoPageでスライドが自動追加されるためslidは最初のまま（描画後は直接制御不可）
            y: CONTINUATION_TOP + headerH + lastPageRows * rowH + 0.32,
            slideCount: Math.min(newSlideCount, MAX_SLIDES_PER_CONTENT),
        };
    }

    return { ...ctx, y: ctx.y + totalTableH + 0.32 };
}

// ============================================================
//  重要事項 ── display_type 分岐 + インラインタグ装飾
//  text_block: カード背景なし。左縦線 + 見出し + 箇条書き（スッキリ）
//  card_alert: テーマ色の薄い背景カード + 左バー + 箇条書き（強調）
// ============================================================
function renderImportantNotes(
    pptx: PptxGenJS,
    ctx: SlideContext,
    notes: ImportantNote[],
    info: BasicInfo,
    num: number,
    total: number
): SlideContext {
    if (!notes || notes.length === 0) return ctx;

    const sanitized = notes.filter(n => n.heading.trim().length > 0 && n.bullets.length > 0);
    if (sanitized.length === 0) return ctx;

    let currentCtx = { ...ctx };

    for (let i = 0; i < sanitized.length; i++) {
        const note = sanitized[i];
        const isCard = note.display_type === "card_alert";

        // 高さ計算
        const headingH = 0.30;
        const bulletLineH = 0.22;
        const cardPadding = isCard ? 0.12 : 0.06;
        const totalBulletH = note.bullets.reduce((sum, b) => {
            return sum + (b.indent_level === 1 ? bulletLineH * 0.9 : bulletLineH);
        }, 0);
        const blockH = cardPadding + headingH + totalBulletH + cardPadding * 0.5;

        // 改ページチェック
        if (currentCtx.y + blockH > PAGE_BOTTOM) {
            if (currentCtx.slideCount < MAX_SLIDES_PER_CONTENT) {
                currentCtx = addContinuationSlide(pptx, info, "notes", num, total, currentCtx);
            } else {
                break;
            }
        }

        if (isCard) {
            // ---- card_alert: テーマ色の薄い背景カード ----
            const theme = ALERT_THEMES[note.card_theme || "info"];

            // 薄い背景色の角丸カード
            currentCtx.slide.addShape("roundRect" as PptxGenJS.ShapeType, {
                x: MX, y: currentCtx.y, w: CW, h: blockH,
                fill: { color: theme.bg },
                rectRadius: 0.1,
                line: { color: theme.accent, width: 0.5 },
            });

            // 左アクセントバー（テーマ色の濃い色）
            currentCtx.slide.addShape("rect" as PptxGenJS.ShapeType, {
                x: MX, y: currentCtx.y, w: 0.06, h: blockH,
                fill: { color: theme.accent },
            });

            // 見出し（テーマ色の濃いテキスト + 太字 + 大きめ）
            currentCtx.slide.addText(note.heading, {
                x: MX + 0.18, y: currentCtx.y + cardPadding * 0.5, w: CW - 0.30, h: headingH,
                fontSize: 14, fontFace: FONT, color: theme.text, bold: true,
                valign: "middle",
            });

            // 箇条書き（インラインタグ装飾 + テーマ色テキスト）
            let bulletY = currentCtx.y + cardPadding * 0.5 + headingH + 0.02;
            for (const bullet of note.bullets) {
                const isIndented = bullet.indent_level === 1;
                const lineH = isIndented ? bulletLineH * 0.9 : bulletLineH;
                const prefix = isIndented ? "  – " : "▸ ";
                const xOffset = isIndented ? 0.34 : 0.18;
                const fontSize = isIndented ? MIN_BODY_FONT - 1 : MIN_BODY_FONT;
                const baseColor = isIndented ? C.textSub : C.textMain;

                const tagParts = parseInlineTags(`${prefix}${bullet.text}`, fontSize, baseColor);
                currentCtx.slide.addText(tagParts, {
                    x: MX + xOffset, y: bulletY,
                    w: CW - xOffset - 0.14, h: lineH,
                    valign: "top",
                });
                bulletY += lineH;
            }
        } else {
            // ---- text_block: カード背景なし。左縦線 + テキストのみ ----
            const sectionColor = SECTION_ACCENTS[i % SECTION_ACCENTS.length];

            // 左縦線（title-bar風）
            currentCtx.slide.addShape("rect" as PptxGenJS.ShapeType, {
                x: MX, y: currentCtx.y + cardPadding * 0.5, w: 0.05, h: headingH,
                fill: { color: sectionColor },
            });

            // 見出し（大きめ + 太字 + アクセント色）
            currentCtx.slide.addText(note.heading, {
                x: MX + 0.14, y: currentCtx.y + cardPadding * 0.5, w: CW - 0.20, h: headingH,
                fontSize: 16, fontFace: FONT, color: C.baseDark, bold: true,
                valign: "middle",
            });

            // 薄い下線（見出しとの区切り）
            currentCtx.slide.addShape("rect" as PptxGenJS.ShapeType, {
                x: MX + 0.14, y: currentCtx.y + cardPadding * 0.5 + headingH,
                w: CW - 0.20, h: 0.01,
                fill: { color: C.gray200 },
            });

            // 箇条書き（インラインタグ装飾 + 濃淡のあるグレーテキスト）
            let bulletY = currentCtx.y + cardPadding * 0.5 + headingH + 0.05;
            for (const bullet of note.bullets) {
                const isIndented = bullet.indent_level === 1;
                const lineH = isIndented ? bulletLineH * 0.9 : bulletLineH;
                const prefix = isIndented ? "  – " : "• ";
                const xOffset = isIndented ? 0.30 : 0.14;
                const fontSize = isIndented ? MIN_BODY_FONT - 1 : MIN_BODY_FONT;
                const baseColor = isIndented ? C.textSub : C.textMain;

                const tagParts = parseInlineTags(`${prefix}${bullet.text}`, fontSize, baseColor);
                currentCtx.slide.addText(tagParts, {
                    x: MX + xOffset, y: bulletY,
                    w: CW - xOffset - 0.14, h: lineH,
                    valign: "top",
                });
                bulletY += lineH;
            }
        }

        currentCtx.y += blockH + 0.10;
    }

    return currentCtx;
}

// ============================================================
//  クリティカルアラート ── 階層付き箇条書き描画
// ============================================================
function renderCriticalAlerts(
    pptx: PptxGenJS,
    ctx: SlideContext,
    alerts: CriticalAlert[],
    info: BasicInfo,
    num: number,
    total: number
): SlideContext {
    if (!alerts || alerts.length === 0) return ctx;

    const sanitized = alerts.filter(a => a.action.trim().length > 0 && a.bullets.length > 0);
    if (sanitized.length === 0) return ctx;

    let currentCtx = { ...ctx };

    // title-bar風セクション見出し描画
    const drawSectionLabel = (targetSlide: PptxGenJS.Slide, labelY: number) => {
        // 左縦線（title-bar風）
        targetSlide.addShape("rect" as PptxGenJS.ShapeType, {
            x: MX, y: labelY, w: 0.05, h: 0.28,
            fill: { color: C.accentRed },
        });
        targetSlide.addText("⚠ 重要アラート", {
            x: MX + 0.12, y: labelY, w: 2, h: 0.28,
            fontSize: 11, fontFace: FONT, color: C.accentRed, bold: true,
            valign: "middle",
        });
    };

    // 最初のセクションラベル
    drawSectionLabel(currentCtx.slide, currentCtx.y);
    currentCtx.y += 0.34;

    for (let i = 0; i < sanitized.length; i++) {
        const alert = sanitized[i];

        // アラートカードの高さを計算
        const actionH = 0.28;
        const bulletLineH = 0.2;
        const totalBulletH = alert.bullets.reduce((sum, b) => {
            return sum + (b.indent_level === 1 ? bulletLineH * 0.9 : bulletLineH);
        }, 0);
        const alertH = actionH + totalBulletH + 0.12;

        // 改ページチェック
        if (currentCtx.y + alertH > PAGE_BOTTOM) {
            if (currentCtx.slideCount < MAX_SLIDES_PER_CONTENT) {
                currentCtx = addContinuationSlide(pptx, info, "notes", num, total, currentCtx);
                drawSectionLabel(currentCtx.slide, currentCtx.y);
                currentCtx.y += 0.34;
            } else {
                break;
            }
        }

        // カード背景（薄赤 + 細い赤枠線 + 角丸 = モダン警告パネル）
        currentCtx.slide.addShape("roundRect" as PptxGenJS.ShapeType, {
            x: MX, y: currentCtx.y, w: CW, h: alertH,
            fill: { color: C.bgCardDanger },
            rectRadius: 0.08,
            line: { color: "FCA5A5", width: 0.5 },
        });

        // 左アクセントバー（赤で警告を強調）
        currentCtx.slide.addShape("rect" as PptxGenJS.ShapeType, {
            x: MX, y: currentCtx.y, w: 0.05, h: alertH,
            fill: { color: C.accentRed },
        });

        // アクション名（赤文字・太字・14pt・十分な余白）
        currentCtx.slide.addText(`⚠ ${alert.action}`, {
            x: MX + 0.18, y: currentCtx.y + 0.10, w: CW - 0.36, h: actionH,
            fontSize: 14, fontFace: FONT, color: C.accentRed, bold: true,
            valign: "middle",
        });

        // 階層付き箇条書き（本文色テキスト・読みやすい余白）
        let bulletY = currentCtx.y + 0.12 + actionH;
        for (const bullet of alert.bullets) {
            const isIndented = bullet.indent_level === 1;
            const lineH = isIndented ? bulletLineH * 0.9 : bulletLineH;
            const prefix = isIndented ? "  – " : "▸ ";
            const xOffset = isIndented ? 0.36 : 0.18;
            const fontSize = isIndented ? MIN_BODY_FONT - 1 : MIN_BODY_FONT;
            const baseColor = isIndented ? C.textSub : C.textMain;

            const tagParts = parseInlineTags(`${prefix}${bullet.text}`, fontSize, baseColor);
            currentCtx.slide.addText(tagParts, {
                x: MX + xOffset, y: bulletY,
                w: CW - xOffset - 0.18, h: lineH,
                valign: "top",
            });
            bulletY += lineH;
        }

        currentCtx.y += alertH + 0.12;
    }

    return currentCtx;
}

// ============================================================
//  レイアウト画像の動的配置
//  ── 現在のスライドの空きスペースに配置。足りなければ次のスライドに。
//  ── 専用スライドは作らず、既存スライドの下部にねじ込む。
// ============================================================
function placeLayoutImage(
    pptx: PptxGenJS,
    ctx: SlideContext,
    base64Image: string,
    info: BasicInfo,
    num: number,
    total: number
) {
    const minImageH = 2.0; // 画像に最低限必要な高さ
    const labelH = 0.3;    // 「レイアウト・動線図」ラベルの高さ
    const neededH = minImageH + labelH + 0.1;

    let targetSlide = ctx.slide;
    let startY = ctx.y + 0.1;

    const remainingSpace = PAGE_BOTTOM - startY;

    // 現在のスライドに十分なスペースがあるか
    if (remainingSpace < neededH) {
        // スペース不足 → 新しいスライドに配置（3枚制限チェック）
        if (ctx.slideCount >= MAX_SLIDES_PER_CONTENT) {
            // 3枚制限のため配置を断念
            return;
        }
        const newCtx = addContinuationSlide(pptx, info, "timeline", num, total, ctx);
        targetSlide = newCtx.slide;
        startY = CONTINUATION_TOP;
    }

    // セクションラベル
    targetSlide.addShape("roundRect" as PptxGenJS.ShapeType, {
        x: MX, y: startY, w: 2.0, h: 0.24,
        fill: { color: C.bgCardInfo }, rectRadius: 0.04,
    });
    targetSlide.addText("🗺️ レイアウト・動線図", {
        x: MX, y: startY, w: 2.0, h: 0.24,
        fontSize: 8, fontFace: FONT, color: C.baseBlue, bold: true,
        align: "center", valign: "middle",
    });

    // 画像配置領域
    const imgTop = startY + labelH;
    const imgBottom = PAGE_BOTTOM - 0.1;
    const availW = CW;
    const availH = imgBottom - imgTop;

    // アスペクト比4:3でfit-contain
    const assumedAspect = 4 / 3;
    let drawW: number;
    let drawH: number;

    if (availW / availH > assumedAspect) {
        drawH = availH;
        drawW = drawH * assumedAspect;
    } else {
        drawW = availW;
        drawH = drawW / assumedAspect;
    }

    // センタリング
    const drawX = MX + (availW - drawW) / 2;
    const drawY = imgTop + (availH - drawH) / 2;

    // カード影
    targetSlide.addShape(pptx.ShapeType.roundRect, {
        x: drawX - 0.04, y: drawY - 0.04,
        w: drawW + 0.08, h: drawH + 0.08,
        fill: { color: C.white },
        shadow: { type: "outer", blur: 4, offset: 1, color: "000000", opacity: 0.06 },
        line: { color: C.gray200, width: 0.5 },
        rectRadius: 0.06,
    });

    // レイアウト画像
    targetSlide.addImage({
        data: `image/png;base64,${base64Image}`,
        x: drawX, y: drawY, w: drawW, h: drawH,
    });
}
