# ============================================
# Google Cloud Run 用 マルチステージビルド Dockerfile
# ============================================

# --- Stage 1: 依存関係インストール ---
FROM node:20-alpine AS deps
RUN apk add --no-cache libc6-compat
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --ignore-scripts

# --- Stage 2: ビルド ---
FROM node:20-alpine AS builder
WORKDIR /app
COPY --from=deps /app/node_modules ./node_modules
COPY . .

# Next.js のテレメトリを無効化
ENV NEXT_TELEMETRY_DISABLED=1

RUN npm run build

# --- Stage 3: 本番用ランナー（最小構成） ---
FROM node:20-alpine AS runner
WORKDIR /app

ENV NODE_ENV=production
ENV NEXT_TELEMETRY_DISABLED=1

# Cloud Run はポート 8080 を使用
ENV PORT=8080
EXPOSE 8080

# セキュリティ: 非rootユーザーで実行
RUN addgroup --system --gid 1001 nodejs
RUN adduser --system --uid 1001 nextjs

# public ディレクトリをコピー（pdf.worker.min.mjs など）
COPY --from=builder /app/public ./public

# standalone 出力をコピー
COPY --from=builder --chown=nextjs:nodejs /app/.next/standalone ./
COPY --from=builder --chown=nextjs:nodejs /app/.next/static ./.next/static

USER nextjs

# Next.js standalone の server.js を起動
# HOSTNAME を 0.0.0.0 に設定して外部からのアクセスを許可
CMD ["node", "server.js"]
