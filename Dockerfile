# vba-lint-mcp — Multi-stage Docker build
# Derived from Rubberduck VBA — GPL-3.0
# Full source included per GPL-3.0 Section 6

# --- Stage 1: Build ---
FROM node:22-alpine AS builder
WORKDIR /app

# Install dependencies first (layer caching)
COPY package*.json ./
RUN npm ci

# Copy source and grammar
COPY . .

# Generate ANTLR4 parser (requires Java for antlr4ng-cli)
RUN apk add --no-cache openjdk17-jre-headless && npm run generate-parser

# Compile TypeScript
RUN npm run build

# --- Stage 2: Runtime ---
FROM node:22-alpine AS runtime
WORKDIR /app

# Run as non-root user
USER node

# Copy only what's needed at runtime
COPY --from=builder --chown=node:node /app/dist ./dist
COPY --from=builder --chown=node:node /app/node_modules ./node_modules
COPY --from=builder --chown=node:node /app/package.json ./

# Include source for GPL-3.0 compliance (Section 6)
COPY --from=builder --chown=node:node /app/src ./src
COPY --from=builder --chown=node:node /app/grammar ./grammar
COPY --from=builder --chown=node:node /app/LICENSE ./
COPY --from=builder --chown=node:node /app/ATTRIBUTION.md ./

# No HEALTHCHECK — stdio transport has no HTTP endpoint.
# The MCP server communicates via stdin/stdout JSON-RPC,
# so there is no port to probe.

CMD ["node", "dist/server.js"]
