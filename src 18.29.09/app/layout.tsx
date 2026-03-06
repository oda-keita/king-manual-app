import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
    title: "KING Operation Manual Generator",
    description:
        "Business Contest KING 公式 — イベント設計書をAIで解析し、運営マニュアルスライドを自動生成するツール",
};

export default function RootLayout({
    children,
}: Readonly<{
    children: React.ReactNode;
}>) {
    return (
        <html lang="ja">
            <body>{children}</body>
        </html>
    );
}
