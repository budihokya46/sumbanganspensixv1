FROM node:22-alpine

ENV NODE_ENV=production
WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY . .

ENV PORT=3000
ENV ADMIN_PASSWORD=admin123

EXPOSE 3000

VOLUME ["/app/data"]

CMD ["node", "server.js"]