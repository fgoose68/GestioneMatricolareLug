# Gestione Modelli L

Applicazione web per la generazione di attestati/modelli “L” a partire da:

- un file Excel contenente i dati dei discenti
- un template Word (.docx) con i segnaposto

I documenti generati vengono scaricati in un archivio ZIP e possono essere stampati successivamente.

**Versione attuale:** `Ver.20.Ott2026`

## Avvio tramite Docker

L’applicazione è containerizzata e raggiungibile in locale sulla porta **2025**.

### Requisiti

- Docker
- Docker Compose

### Avvio

```bash
docker compose up --build -d
```

L’app sarà disponibile all’indirizzo:

```text
http://localhost:2025
```

### Arresto

```bash
docker compose down
```

## Pubblicazione con dominio (es. gm.teckapp.ovh)

Per rendere l’app raggiungibile da un dominio, configurare un reverse proxy Nginx che punti a `http://localhost:2025` (o all’IP del server Docker se Nginx è su un’altra macchina).

Esempio di configurazione per Ubuntu in `/etc/nginx/sites-available/gm.teckapp.ovh`:

```nginx
server {
    listen 80;
    server_name gm.teckapp.ovh;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl http2;
    server_name gm.teckapp.ovh;

    ssl_certificate     /etc/letsencrypt/live/gm.teckapp.ovh/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/gm.teckapp.ovh/privkey.pem;

    location / {
        proxy_pass http://127.0.0.1:2025;
        proxy_http_version 1.1;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

Abilitare il sito:

```bash
sudo ln -s /etc/nginx/sites-available/gm.teckapp.ovh /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl reload nginx
```

Ottenere il certificato SSL con Let’s Encrypt:

```bash
sudo certbot --nginx -d gm.teckapp.ovh
```

Infine, assicurarsi che il record DNS punti all’IP del server Nginx:

```text
gm.teckapp.ovh  A  <IP_SERVER>
```

## Stack tecnologico

- React 18 + TypeScript
- Vite
- Tailwind CSS + shadcn/ui
- docxtemplater + pizzip per i documenti Word
- xlsx per la lettura dei file Excel
- Nginx per il serve dei file statici in produzione

## Note

- La porta interna del container è la `80` (Nginx); sulla macchina host viene mappata sulla `2025`.
- L’app è una SPA (Single Page Application), quindi la configurazione Nginx deve includere il fallback su `index.html` per il routing React.
