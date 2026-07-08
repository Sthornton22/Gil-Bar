# Gil-Bar Proposal Engine — Render deployment
FROM node:20-slim

WORKDIR /app

# install dependencies first (better build caching)
COPY package.json ./
RUN npm install --omit=dev

# app code + assets
COPY . .

# Render injects PORT; server.js reads process.env.PORT
EXPOSE 3000
CMD ["npm", "start"]
