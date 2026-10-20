# Stage 1: Build dell'app
FROM node:22-alpine AS builder

WORKDIR /app

# Abilita pnpm tramite corepack (la versione è definita in package.json)
RUN corepack enable && corepack prepare pnpm@10.10.0 --activate

# Copia solo i file di dipendenza per sfruttare la cache Docker
COPY package.json pnpm-lock.yaml ./
RUN pnpm install --frozen-lockfile

# Copia il resto del codice sorgente e builda
COPY . .
RUN pnpm build

# Stage 2: Serve i file statici con Nginx
FROM nginx:alpine

# Copia la configurazione Nginx
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Copia l'output della build di Vite nella cartella servita da Nginx
COPY --from=builder /app/dist /usr/share/nginx/html

EXPOSE 80

CMD ["nginx", "-g", "daemon off;"]
