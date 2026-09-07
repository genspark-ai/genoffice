# GenOffice Web Server Dockerfile
# Multi-stage build for smaller image

FROM node:22-alpine AS builder

WORKDIR /app

# Copy workspace files
COPY package*.json ./
COPY apps/web-server/package*.json ./

# Install dependencies
RUN npm ci --ignore-scripts

# Build TypeScript
RUN npm run build -w @genoffice/web-server

# Production stage
FROM node:22-alpine

WORKDIR /app

# Create non-root user
RUN addgroup -g 1001 -S nodejs && \
    adduser -S nodejs -u 1001

# Copy built files
COPY --from=builder /app/apps/web-server/dist ./dist
COPY --from=builder /app/node_modules ./node_modules

# Set ownership
RUN chown -R nodejs:nodejs /app

# Switch to non-root user
USER nodejs

# Environment defaults
ENV HOST=0.0.0.0
ENV PORT=8080
ENV NODE_ENV=production

# Expose port
EXPOSE 8080

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD wget --no-verbose --tries=1 --spider http://localhost:8080/health || exit 1

# Start command
CMD ["node", "dist/index.js"]
