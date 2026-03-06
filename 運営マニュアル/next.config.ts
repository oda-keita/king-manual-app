import type { NextConfig } from "next";

const nextConfig: NextConfig = {
    // Cloud Run（Docker）用スタンドアロンビルド
    output: "standalone",
    // サーバーサイドでのみ使用するパッケージ
    serverExternalPackages: ["pdf-parse"],
};

export default nextConfig;
