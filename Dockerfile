# Nextcloud Excel Sync - Frontend & Express Server Dockerfile for Coolify
FROM node:20-slim AS builder

WORKDIR /app

# Install build dependencies
COPY package.json package-lock.json* ./
RUN npm install

# Copy source files
COPY . .

# Build Vite frontend assets and bundle Express backend to dist/server.cjs
RUN npm run build

# Production Runner
FROM node:20-slim AS runner

WORKDIR /app

# Install curl for healthcheck
RUN apt-get update && apt-get install -y --no-install-recommends curl \
    && rm -rf /var/lib/apt/lists/*

ENV NODE_ENV=production
ENV PORT=3000

# Copy package configs and install production dependencies only
COPY package.json package-lock.json* ./
RUN npm install --omit=dev

# Copy compiled frontend and bundled server from builder stage
COPY --from=builder /app/dist ./dist
COPY --from=builder /app/public ./public

# Healthcheck for Coolify
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD curl -f http://localhost:3000/health || exit 1

EXPOSE 3000

CMD ["node", "dist/server.cjs"]
