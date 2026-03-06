// 認証API: パスワードをサーバーサイドで検証する
import { NextRequest, NextResponse } from "next/server";

export async function POST(request: NextRequest) {
    try {
        const { password } = await request.json();
        const appPassword = process.env.APP_PASSWORD;

        if (!appPassword) {
            console.error("[Auth] APP_PASSWORD 環境変数が設定されていません");
            return NextResponse.json(
                { error: "サーバー設定エラー" },
                { status: 500 }
            );
        }

        if (password === appPassword) {
            return NextResponse.json({ success: true });
        } else {
            return NextResponse.json(
                { success: false, error: "パスワードが正しくありません" },
                { status: 401 }
            );
        }
    } catch {
        return NextResponse.json(
            { error: "リクエストの解析に失敗しました" },
            { status: 400 }
        );
    }
}
