FROM node:20-bookworm-slim

# LibreOffice (headless) for .docx -> PDF, plus Liberation fonts (metric-compatible with Arial)
# so the "Arial" font in the helper library renders correctly.
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice-core libreoffice-writer \
      fonts-liberation fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

ENV HOME=/tmp
WORKDIR /app

COPY package.json ./
RUN npm install --omit=dev

COPY . .

EXPOSE 8088
CMD ["node", "server.js"]
