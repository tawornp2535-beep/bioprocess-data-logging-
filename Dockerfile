FROM node:18-alpine AS build
WORKDIR /app

# Install dependencies (including dev deps for Vite build)
COPY package.json package-lock.json* ./
RUN npm ci --silent

# Copy source and build frontend
COPY . .
RUN npm run build

FROM node:18-alpine AS prod
WORKDIR /app

# Copy only production deps
COPY package.json package-lock.json* ./
RUN npm ci --production --silent || true

# Copy app files including built frontend
COPY --from=build /app .

ENV NODE_ENV=production
EXPOSE 5000
CMD ["node", "server.js"]
