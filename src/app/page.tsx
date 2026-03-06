"use client";

import { useState, useRef, useCallback, useEffect } from "react";
import { pdfToBase64Images } from "@/lib/pdfToImages";
import {
    FileText,
    Image,
    LayoutGrid,
    Users,
    Upload,
    Sparkles,
    CheckCircle2,
    AlertCircle,
    Download,
    X,
    Loader2,
    ArrowRight,
    FileUp,
    Cpu,
    Presentation,
    RefreshCw,
    Lightbulb,
    Lock,
    ShieldCheck,
} from "lucide-react";



// 処理のステップ定義
type Step = "idle" | "uploading" | "analyzing" | "generating" | "done";
type ErrorType = "file_parse" | "ai_analysis" | "slide_generation" | "validation" | "unknown";

interface ErrorInfo {
    message: string;
    errorType: ErrorType;
    fileName?: string;
}

interface StepInfo {
    key: Step;
    label: string;
    activeLabel: string;
    icon: React.ReactNode;
}

const STEPS: StepInfo[] = [
    {
        key: "uploading",
        label: "ファイル読み込み完了",
        activeLabel: "ファイルを読み込み中...",
        icon: <FileUp size={14} />,
    },
    {
        key: "analyzing",
        label: "AI解析完了",
        activeLabel: "AIが設計書を解析中...",
        icon: <Cpu size={14} />,
    },
    {
        key: "generating",
        label: "スライド生成完了",
        activeLabel: "スライドをレンダリング中...",
        icon: <Presentation size={14} />,
    },
];

function formatFileSize(bytes: number): string {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

// Text Ticker — スロットマシーン風に切り替わるステータステキスト
function useLoadingText(isActive: boolean, step: Step): { text: string; key: number } {
    const messages: Record<string, string[]> = {
        uploading: ["ファイルを読み込んでいます..."],
        analyzing: [
            "AIが設計書を解析中...",
            "イベント構造を分析中...",
            "タイムラインを抽出中...",
            "重要事項を整理中...",
        ],
        generating: [
            "スライドをレンダリング中...",
            "レイアウトを最適化中...",
            "PowerPointを構築中...",
        ],
    };
    const [index, setIndex] = useState(0);
    const [text, setText] = useState("");
    const [tickerKey, setTickerKey] = useState(0);

    useEffect(() => {
        if (!isActive || step === "idle" || step === "done") return;
        const list = messages[step] || [];
        if (list.length === 0) return;
        setText(list[0]);
        setIndex(0);
        setTickerKey((k) => k + 1);
        if (list.length <= 1) return;

        const interval = setInterval(() => {
            setIndex((prev) => {
                const next = (prev + 1) % list.length;
                setText(list[next]);
                setTickerKey((k) => k + 1);
                return next;
            });
        }, 3000);
        return () => clearInterval(interval);
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [isActive, step]);

    return { text, key: tickerKey };
}

// Stagger表示 — 各要素を順次フェードイン
function useStagger(isActive: boolean, count: number, baseDelay = 80): boolean[] {
    const [visible, setVisible] = useState<boolean[]>(Array(count).fill(false));
    useEffect(() => {
        if (!isActive) {
            setVisible(Array(count).fill(false));
            return;
        }
        const timers: ReturnType<typeof setTimeout>[] = [];
        for (let i = 0; i < count; i++) {
            timers.push(
                setTimeout(() => {
                    setVisible((prev) => {
                        const next = [...prev];
                        next[i] = true;
                        return next;
                    });
                }, baseDelay * i)
            );
        }
        return () => timers.forEach(clearTimeout);
    }, [isActive, count, baseDelay]);
    return visible;
}

export default function Home() {
    // === 認証State ===
    const [isAuthenticated, setIsAuthenticated] = useState(false);
    const [password, setPassword] = useState("");
    const [loginError, setLoginError] = useState(false);
    const [isShaking, setIsShaking] = useState(false);
    const [showMain, setShowMain] = useState(false);
    const [loginExiting, setLoginExiting] = useState(false);
    const passwordInputRef = useRef<HTMLInputElement>(null);

    // === メインアプリState ===
    const [files, setFiles] = useState<File[]>([]);
    const [layoutFile, setLayoutFile] = useState<File | null>(null);
    const [timelineImage, setTimelineImage] = useState<File | null>(null);
    const [roleImage, setRoleImage] = useState<File | null>(null);
    const [currentStep, setCurrentStep] = useState<Step>("idle");
    const [errorInfo, setErrorInfo] = useState<ErrorInfo | null>(null);
    const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
    const [isDragOver, setIsDragOver] = useState(false);
    const fileInputRef = useRef<HTMLInputElement>(null);
    const layoutInputRef = useRef<HTMLInputElement>(null);
    const timelineImageRef = useRef<HTMLInputElement>(null);
    const roleImageRef = useRef<HTMLInputElement>(null);

    const isProcessing = currentStep !== "idle" && currentStep !== "done";
    const loadingTicker = useLoadingText(isProcessing, currentStep);

    // === Stagger ===
    const loginStagger = useStagger(!isAuthenticated, 6, 100);
    const mainStagger = useStagger(showMain, 8, 90);

    // === ログイン処理 — シネマティック遷移 ===
    const handleLogin = useCallback(async () => {
        try {
            const res = await fetch("/api/auth", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ password }),
            });
            const data = await res.json();
            if (data.success) {
                setLoginError(false);
                setLoginExiting(true);
                // シネマティック: 光の爆発 → カードが消える → メイン画面出現
                setTimeout(() => {
                    setIsAuthenticated(true);
                    setTimeout(() => setShowMain(true), 100);
                }, 800);
            } else {
                setLoginError(true);
                setIsShaking(true);
                setTimeout(() => setIsShaking(false), 500);
                passwordInputRef.current?.focus();
            }
        } catch {
            setLoginError(true);
            setIsShaking(true);
            setTimeout(() => setIsShaking(false), 500);
            passwordInputRef.current?.focus();
        }
    }, [password]);

    const handlePasswordKeyDown = useCallback(
        (e: React.KeyboardEvent) => {
            if (e.key === "Enter") {
                handleLogin();
            }
        },
        [handleLogin]
    );

    // 画像ファイル → Base64変換ヘルパー
    const fileToBase64 = (file: File): Promise<string> => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => {
                const result = reader.result as string;
                const base64 = result.split(",")[1] || result;
                resolve(base64);
            };
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
    };

    // エラーカテゴリに応じたUI情報を返す
    const getErrorDisplay = (info: ErrorInfo) => {
        switch (info.errorType) {
            case "file_parse":
                return {
                    title: "ファイル読み取りエラー",
                    hint: "ファイルが破損していないか、パスワード保護されていないかご確認ください。",
                };
            case "ai_analysis":
                return {
                    title: "AI解析エラー",
                    hint: "AIのレスポンスに問題がありました。もう一度お試しください。",
                };
            case "slide_generation":
                return {
                    title: "スライド生成エラー",
                    hint: "スライドの描画中に問題が発生しました。ファイル内容を変更して再度お試しください。",
                };
            case "validation":
                return {
                    title: "入力エラー",
                    hint: "ファイル形式やサイズをご確認ください。",
                };
            default:
                return {
                    title: "予期しないエラー",
                    hint: "時間をおいてから再度お試しください。問題が続く場合は管理者にご連絡ください。",
                };
        }
    };

    // ファイル追加ハンドラ（複数対応）
    const handleFilesAdd = useCallback(
        (newFiles: FileList | File[]) => {
            setErrorInfo(null);
            setDownloadUrl(null);
            setCurrentStep("idle");

            const validFiles: File[] = [];
            for (const file of Array.from(newFiles)) {
                const ext = file.name.toLowerCase().split(".").pop();
                if (!ext || !["pdf", "docx"].includes(ext)) {
                    setErrorInfo({ message: `「${file.name}」は対応していません（.pdf / .docx のみ）`, errorType: "validation", fileName: file.name });
                    continue;
                }
                if (file.size > 10 * 1024 * 1024) {
                    setErrorInfo({ message: `「${file.name}」のサイズが大きすぎます（上限: 10MB）`, errorType: "validation", fileName: file.name });
                    continue;
                }
                if (files.some((f) => f.name === file.name && f.size === file.size)) {
                    continue;
                }
                validFiles.push(file);
            }

            if (validFiles.length > 0) {
                setFiles((prev) => [...prev, ...validFiles]);
            }
        },
        [files]
    );

    // ドラッグ&ドロップ
    const handleDragOver = useCallback((e: React.DragEvent) => {
        e.preventDefault();
        setIsDragOver(true);
    }, []);

    const handleDragLeave = useCallback((e: React.DragEvent) => {
        e.preventDefault();
        setIsDragOver(false);
    }, []);

    const handleDrop = useCallback(
        (e: React.DragEvent) => {
            e.preventDefault();
            setIsDragOver(false);
            if (e.dataTransfer.files.length > 0) {
                handleFilesAdd(e.dataTransfer.files);
            }
        },
        [handleFilesAdd]
    );

    // ファイル入力change
    const handleInputChange = useCallback(
        (e: React.ChangeEvent<HTMLInputElement>) => {
            if (e.target.files && e.target.files.length > 0) {
                handleFilesAdd(e.target.files);
            }
            if (fileInputRef.current) {
                fileInputRef.current.value = "";
            }
        },
        [handleFilesAdd]
    );

    // 個別ファイル除去
    const handleRemoveFile = useCallback((index: number) => {
        setFiles((prev) => prev.filter((_, i) => i !== index));
        setErrorInfo(null);
    }, []);

    // スライド生成実行
    const handleGenerate = useCallback(async () => {
        if (files.length === 0) return;

        setErrorInfo(null);
        setDownloadUrl(null);

        try {
            setCurrentStep("uploading");
            const formData = new FormData();
            files.forEach((f) => formData.append("files", f));

            // レイアウト図PDFがある場合、ブラウザ側で画像化してBase64をFormDataに追加
            if (layoutFile) {
                console.log(`[SlideGen] レイアウト図PDF「${layoutFile.name}」を画像化中...`);
                try {
                    const layoutImages = await pdfToBase64Images(layoutFile);
                    console.log(`[SlideGen] レイアウト図: ${layoutImages.length}ページを画像化完了`);
                    layoutImages.forEach((img) => formData.append("layoutImages", img));
                } catch (imgErr) {
                    console.error("[SlideGen] レイアウト図の画像化に失敗:", imgErr);
                    setErrorInfo({
                        message: `レイアウト図の画像化に失敗しました: ${imgErr instanceof Error ? imgErr.message : String(imgErr)}`,
                        errorType: "file_parse",
                        fileName: layoutFile.name,
                    });
                    setCurrentStep("idle");
                    return;
                }
            }

            // 全体TL用画像のBase64化
            if (timelineImage) {
                const b64 = await fileToBase64(timelineImage);
                formData.append("timelineImage", b64);
            }
            // 役職割用画像のBase64化
            if (roleImage) {
                const b64 = await fileToBase64(roleImage);
                formData.append("roleImage", b64);
            }

            await new Promise((r) => setTimeout(r, 400));
            setCurrentStep("analyzing");

            console.log(`[SlideGen] ${files.length}ファイルをAPIに送信中...`);
            const response = await fetch("/api/generate", {
                method: "POST",
                body: formData,
            });
            console.log("[SlideGen] レスポンス:", response.status);

            if (!response.ok) {
                let errorMessage = `サーバーエラー（${response.status}）`;
                let errorType: ErrorType = "unknown";
                let fileName: string | undefined;
                try {
                    const errorData = await response.json();
                    errorMessage = errorData.error || errorMessage;
                    errorType = errorData.errorType || "unknown";
                    fileName = errorData.fileName;
                } catch {
                    try {
                        const errorText = await response.text();
                        if (errorText) errorMessage = errorText;
                    } catch {
                        // noop
                    }
                }
                throw { message: errorMessage, errorType, fileName };
            }

            setCurrentStep("generating");
            const blob = await response.blob();
            console.log("[SlideGen] Blob:", blob.size, "bytes");

            if (blob.size === 0) {
                throw new Error("生成されたファイルが空です。もう一度お試しください。");
            }

            const url = URL.createObjectURL(blob);
            setDownloadUrl(url);
            setCurrentStep("done");

            // 自動ダウンロード
            const a = document.createElement("a");
            a.href = url;
            a.download = "event-slides.pptx";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        } catch (err: unknown) {
            console.error("[SlideGen] エラー:", err);
            if (
                err &&
                typeof err === "object" &&
                "errorType" in err
            ) {
                const typedErr = err as ErrorInfo;
                setErrorInfo({
                    message: typedErr.message || "予期しないエラーが発生しました。",
                    errorType: typedErr.errorType || "unknown",
                    fileName: typedErr.fileName,
                });
            } else {
                setErrorInfo({
                    message:
                        err instanceof Error
                            ? err.message
                            : "予期しないエラーが発生しました。",
                    errorType: "unknown",
                });
            }
            setCurrentStep("idle");
        }
    }, [files, layoutFile, timelineImage, roleImage]);

    // リセット
    const handleReset = useCallback(() => {
        if (downloadUrl) URL.revokeObjectURL(downloadUrl);
        setFiles([]);
        setLayoutFile(null);
        setTimelineImage(null);
        setRoleImage(null);
        setCurrentStep("idle");
        setErrorInfo(null);
        setDownloadUrl(null);
        if (fileInputRef.current) fileInputRef.current.value = "";
        if (layoutInputRef.current) layoutInputRef.current.value = "";
        if (timelineImageRef.current) timelineImageRef.current.value = "";
        if (roleImageRef.current) roleImageRef.current.value = "";
    }, [downloadUrl]);

    const getStepStatus = (stepKey: Step): "done" | "active" | "pending" => {
        const stepOrder: Step[] = ["uploading", "analyzing", "generating", "done"];
        const currentIndex = stepOrder.indexOf(currentStep);
        const stepIndex = stepOrder.indexOf(stepKey);
        if (currentStep === "idle") return "pending";
        if (stepIndex < currentIndex) return "done";
        if (stepIndex === currentIndex) return "active";
        return "pending";
    };

    const getProgressPercent = (): number => {
        switch (currentStep) {
            case "uploading": return 20;
            case "analyzing": return 55;
            case "generating": return 85;
            case "done": return 100;
            default: return 0;
        }
    };

    // ===== ログイン画面 =====
    if (!isAuthenticated) {
        return (
            <div className={`login-screen ${loginExiting ? "login-screen--exit" : ""}`}>
                <div className="login-screen__glow" />
                <div className={`login-card ${isShaking ? "login-card--shake" : ""}`}>
                    {/* KINGロゴ */}
                    <div
                        className="login-stagger login-card__logo"
                        style={{ animationDelay: `${0.1}s` }}
                    >
                        <div className="login-card__logo-icon">
                            <ShieldCheck size={28} />
                        </div>
                    </div>
                    <h1
                        className="login-stagger login-card__title"
                        style={{ animationDelay: `${0.2}s` }}
                    >
                        Business Contest
                        <br />
                        <span className="login-card__title-accent">KING</span>
                    </h1>
                    <p
                        className="login-stagger login-card__subtitle"
                        style={{ animationDelay: `${0.3}s` }}
                    >
                        システムのロックを解除してください
                    </p>

                    {/* パスワード入力 */}
                    <div
                        className="login-stagger login-card__input-wrapper"
                        style={{ animationDelay: `${0.4}s` }}
                    >
                        <Lock size={16} className="login-card__input-icon" />
                        <input
                            ref={passwordInputRef}
                            type="password"
                            value={password}
                            onChange={(e) => {
                                setPassword(e.target.value);
                                setLoginError(false);
                            }}
                            onKeyDown={handlePasswordKeyDown}
                            placeholder="パスワードを入力"
                            className={`login-card__input ${loginError ? "login-card__input--error" : ""}`}
                            autoFocus
                        />
                    </div>
                    {loginError && (
                        <p className="login-card__error">
                            <AlertCircle size={12} />
                            パスワードが正しくありません
                        </p>
                    )}

                    {/* Enterボタン */}
                    <button
                        onClick={handleLogin}
                        className="login-stagger login-card__btn"
                        style={{ animationDelay: `${0.5}s` }}
                    >
                        <span>Enter</span>
                        <ArrowRight size={16} />
                    </button>

                    <p
                        className="login-stagger login-card__footer"
                        style={{ animationDelay: `${0.6}s` }}
                    >
                        KING Operation Manual Generator
                    </p>
                </div>
            </div>
        );
    }

    return (
        <main className={`main ${showMain ? "main--visible" : "main--hidden"}`}>
            {/* ===== ヒーローヘッダー ===== */}
            <header className="hero">
                <div
                    className={`stagger-item ${mainStagger[0] ? "stagger-item--visible" : ""}`}
                    style={{ transitionDelay: "0s" }}
                >
                    <div className="hero__badge">
                        <span className="hero__badge-dot" />
                        KING OFFICIAL TOOL
                    </div>
                </div>
                <h1
                    className={`hero__title stagger-item ${mainStagger[1] ? "stagger-item--visible" : ""}`}
                    style={{ transitionDelay: "0.08s" }}
                >
                    マニュアル作成を
                    <br />
                    <span className="hero__title-accent">完全自動化</span>
                </h1>
                <p
                    className={`hero__subtitle stagger-item ${mainStagger[2] ? "stagger-item--visible" : ""}`}
                    style={{ transitionDelay: "0.16s" }}
                >
                    Business Contest KING Operation Manual Generator
                </p>
            </header>

            {/* ===== メインコンテンツ ===== */}
            <div className="container">
                {currentStep === "done" ? (
                    /* ===== 完了画面 ===== */
                    <div className="result">
                        <div className="result__icon-container">
                            <CheckCircle2 size={36} />
                        </div>
                        <h2 className="result__title">生成完了</h2>
                        <p className="result__text">
                            統合スライドの生成が完了しました。
                            <br />
                            ファイルは自動ダウンロードされます。
                        </p>
                        {downloadUrl && (
                            <a
                                href={downloadUrl}
                                download="event-slides.pptx"
                                className="download-btn"
                            >
                                <Download size={18} />
                                PowerPointをダウンロード
                            </a>
                        )}
                        <button onClick={handleReset} className="reset-btn">
                            <RefreshCw size={14} style={{ marginRight: '0.4rem' }} />
                            別のファイルを変換する
                        </button>
                    </div>
                ) : (
                    <>
                        {/* ===== アップロードセクション ===== */}
                        <div
                            className={`section-header stagger-item ${mainStagger[3] ? "stagger-item--visible" : ""}`}
                            style={{ transitionDelay: "0.24s" }}
                        >
                            <div className="section-header__icon">
                                <Upload size={18} />
                            </div>
                            <div className="section-header__text">
                                <h2>ファイルをアップロード</h2>
                                <p>設計書と補助素材を選択してください</p>
                            </div>
                        </div>

                        {/* ===== ドロップゾーングリッド ===== */}
                        <div
                            className={`upload-grid stagger-item ${mainStagger[4] ? "stagger-item--visible" : ""}`}
                            style={{ transitionDelay: "0.32s" }}
                        >
                            {/* メインドロップゾーン: 設計書 */}
                            <div
                                className={`dropzone-card dropzone-card--main ${isDragOver ? "dropzone-card--active" : ""} ${files.length > 0 ? "dropzone-card--has-file" : ""}`}
                                onDragOver={handleDragOver}
                                onDragLeave={handleDragLeave}
                                onDrop={handleDrop}
                                onClick={() => !isProcessing && fileInputRef.current?.click()}
                            >
                                <span className="dropzone-card__badge">必須</span>
                                <div className="dropzone-card__icon">
                                    <FileText size={36} strokeWidth={1.5} />
                                </div>
                                <span className="dropzone-card__label">
                                    {files.length > 0
                                        ? `${files.length}件の設計書を選択中`
                                        : "設計書 (PDF / DOCX)"}
                                </span>
                                <span className="dropzone-card__hint">
                                    ドラッグ＆ドロップ または クリックして選択
                                    <br />
                                    複数ファイル対応 ・ 各10MB以下
                                </span>
                                <input
                                    ref={fileInputRef}
                                    type="file"
                                    accept=".pdf,.docx"
                                    multiple
                                    onChange={handleInputChange}
                                    className="dropzone__input"
                                    disabled={isProcessing}
                                />
                            </div>

                            {/* レイアウト図 */}
                            <div
                                className={`dropzone-card ${layoutFile ? "dropzone-card--has-file" : ""}`}
                                onClick={() => !isProcessing && layoutInputRef.current?.click()}
                            >
                                {layoutFile && !isProcessing && (
                                    <button
                                        className="dropzone-card__remove"
                                        onClick={(e) => {
                                            e.stopPropagation();
                                            setLayoutFile(null);
                                            if (layoutInputRef.current) layoutInputRef.current.value = "";
                                        }}
                                        title="除去"
                                    >
                                        <X size={12} />
                                    </button>
                                )}
                                <span className="dropzone-card__badge">任意</span>
                                <div className="dropzone-card__icon">
                                    <LayoutGrid size={28} strokeWidth={1.5} />
                                </div>
                                <span className="dropzone-card__label">レイアウト図</span>
                                {layoutFile ? (
                                    <span className="dropzone-card__filename">{layoutFile.name}</span>
                                ) : (
                                    <span className="dropzone-card__hint">会場配置図 (PDF)</span>
                                )}
                                <input
                                    ref={layoutInputRef}
                                    type="file"
                                    accept=".pdf"
                                    onChange={(e) => {
                                        const f = e.target.files?.[0];
                                        if (f) {
                                            if (f.size > 20 * 1024 * 1024) {
                                                setErrorInfo({ message: `レイアウト図「${f.name}」が大きすぎます（上限: 20MB）`, errorType: "validation", fileName: f.name });
                                                return;
                                            }
                                            setLayoutFile(f);
                                            setErrorInfo(null);
                                        }
                                    }}
                                    className="dropzone__input"
                                    disabled={isProcessing}
                                />
                            </div>

                            {/* 全体TL画像 */}
                            <div
                                className={`dropzone-card ${timelineImage ? "dropzone-card--has-file" : ""}`}
                                onClick={() => !isProcessing && timelineImageRef.current?.click()}
                            >
                                {timelineImage && !isProcessing && (
                                    <button
                                        className="dropzone-card__remove"
                                        onClick={(e) => {
                                            e.stopPropagation();
                                            setTimelineImage(null);
                                            if (timelineImageRef.current) timelineImageRef.current.value = "";
                                        }}
                                        title="除去"
                                    >
                                        <X size={12} />
                                    </button>
                                )}
                                <span className="dropzone-card__badge">任意</span>
                                <div className="dropzone-card__icon">
                                    <Image size={28} strokeWidth={1.5} />
                                </div>
                                <span className="dropzone-card__label">全体TL画像</span>
                                {timelineImage ? (
                                    <span className="dropzone-card__filename">{timelineImage.name}</span>
                                ) : (
                                    <span className="dropzone-card__hint">タイムライン画像</span>
                                )}
                                <input
                                    ref={timelineImageRef}
                                    type="file"
                                    accept="image/*"
                                    onChange={(e) => {
                                        const f = e.target.files?.[0];
                                        if (f) {
                                            setTimelineImage(f);
                                            setErrorInfo(null);
                                        }
                                    }}
                                    className="dropzone__input"
                                    disabled={isProcessing}
                                />
                            </div>

                            {/* 役職割画像 */}
                            <div
                                className={`dropzone-card ${roleImage ? "dropzone-card--has-file" : ""}`}
                                onClick={() => !isProcessing && roleImageRef.current?.click()}
                            >
                                {roleImage && !isProcessing && (
                                    <button
                                        className="dropzone-card__remove"
                                        onClick={(e) => {
                                            e.stopPropagation();
                                            setRoleImage(null);
                                            if (roleImageRef.current) roleImageRef.current.value = "";
                                        }}
                                        title="除去"
                                    >
                                        <X size={12} />
                                    </button>
                                )}
                                <span className="dropzone-card__badge">任意</span>
                                <div className="dropzone-card__icon">
                                    <Users size={28} strokeWidth={1.5} />
                                </div>
                                <span className="dropzone-card__label">役職割画像</span>
                                {roleImage ? (
                                    <span className="dropzone-card__filename">{roleImage.name}</span>
                                ) : (
                                    <span className="dropzone-card__hint">役職割当の画像</span>
                                )}
                                <input
                                    ref={roleImageRef}
                                    type="file"
                                    accept="image/*"
                                    onChange={(e) => {
                                        const f = e.target.files?.[0];
                                        if (f) {
                                            setRoleImage(f);
                                            setErrorInfo(null);
                                        }
                                    }}
                                    className="dropzone__input"
                                    disabled={isProcessing}
                                />
                            </div>
                        </div>

                        {/* ===== ファイルリスト ===== */}
                        {files.length > 0 && (
                            <div
                                className={`file-list stagger-item ${mainStagger[5] ? "stagger-item--visible" : ""}`}
                                style={{ transitionDelay: "0.40s" }}
                            >
                                <div className="file-list__header">
                                    <FileText size={14} style={{ color: 'var(--color-accent)' }} />
                                    <span className="file-list__count">
                                        {files.length}件のファイルを選択中
                                    </span>
                                </div>
                                <div className="file-list__items">
                                    {files.map((file, index) => (
                                        <div key={`${file.name}-${file.size}`} className="file-item">
                                            <div className="file-item__icon">
                                                <FileText size={18} />
                                            </div>
                                            <div className="file-item__details">
                                                <p className="file-item__name">{file.name}</p>
                                                <p className="file-item__size">
                                                    {formatFileSize(file.size)}
                                                </p>
                                            </div>
                                            {!isProcessing && (
                                                <button
                                                    onClick={(e) => {
                                                        e.stopPropagation();
                                                        handleRemoveFile(index);
                                                    }}
                                                    className="file-item__remove"
                                                    title="ファイルを除去"
                                                >
                                                    <X size={14} />
                                                </button>
                                            )}
                                        </div>
                                    ))}
                                </div>
                            </div>
                        )}

                        {/* ===== 生成ボタン ===== */}
                        <div
                            className={`stagger-item ${mainStagger[6] ? "stagger-item--visible" : ""}`}
                            style={{ transitionDelay: "0.48s" }}
                        >
                            <button
                                onClick={handleGenerate}
                                disabled={files.length === 0 || isProcessing}
                                className="generate-btn"
                            >
                                {isProcessing ? (
                                    <>
                                        <Loader2 size={18} className="generate-btn__icon" style={{ animation: 'spin 0.8s linear infinite' }} />
                                        <span className="generate-btn__text">
                                            <span className="text-ticker">
                                                <span key={loadingTicker.key} className="text-ticker__item">
                                                    {loadingTicker.text || "処理中..."}
                                                </span>
                                            </span>
                                        </span>
                                    </>
                                ) : (
                                    <>
                                        <Sparkles size={18} className="generate-btn__icon" />
                                        <span className="generate-btn__text">
                                            統合スライドを生成する
                                        </span>
                                        <ArrowRight size={16} className="generate-btn__icon" />
                                    </>
                                )}
                            </button>
                        </div>

                        {/* ===== 処理進捗表示 ===== */}
                        {isProcessing && (
                            <div className="progress-section">
                                <div className="progress-bar-container">
                                    <div
                                        className="progress-bar"
                                        style={{ width: `${getProgressPercent()}%` }}
                                    />
                                </div>
                                <div className="progress-steps">
                                    {STEPS.map((step) => {
                                        const status = getStepStatus(step.key);
                                        return (
                                            <div
                                                key={step.key}
                                                className={`progress-step progress-step--${status}`}
                                            >
                                                <div className="progress-step__icon">
                                                    {status === "done" ? (
                                                        <CheckCircle2 size={14} />
                                                    ) : status === "active" ? (
                                                        <span className="spinner-ring" />
                                                    ) : (
                                                        step.icon
                                                    )}
                                                </div>
                                                <div className="progress-step__content">
                                                    <span className="progress-step__label">
                                                        {status === "active" ? (
                                                            <span className="text-ticker">
                                                                <span key={`${step.key}-${loadingTicker.key}`} className="text-ticker__item">
                                                                    {step.activeLabel}
                                                                </span>
                                                            </span>
                                                        ) : (
                                                            step.label
                                                        )}
                                                    </span>
                                                </div>
                                            </div>
                                        );
                                    })}
                                </div>
                            </div>
                        )}
                    </>
                )}

                {/* ===== エラー表示 ===== */}
                {errorInfo && (() => {
                    const display = getErrorDisplay(errorInfo);
                    return (
                        <div className="error">
                            <div className="error__header">
                                <AlertCircle size={16} className="error__icon" />
                                <span className="error__title">{display.title}</span>
                            </div>
                            <p className="error__text">{errorInfo.message}</p>
                            <p className="error__hint">
                                <Lightbulb size={12} style={{ marginRight: '0.3rem', verticalAlign: 'middle' }} />
                                {display.hint}
                            </p>
                        </div>
                    );
                })()}
            </div>

            {/* ===== フッター ===== */}
            <footer
                className={`footer stagger-item ${mainStagger[7] ? "stagger-item--visible" : ""}`}
                style={{ transitionDelay: "0.56s" }}
            >
                <p className="footer__text">
                    Business Contest KING — Powered by Gemini AI × PptxGenJS
                </p>
            </footer>
        </main>
    );
}
