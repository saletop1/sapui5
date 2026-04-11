FROM node:20-alpine

WORKDIR /app

# Copy package files untuk install dependency
COPY package*.json ./
RUN npm install

# Copy seluruh kodingan
COPY . .

# Jalankan server.js
EXPOSE 3000
CMD ["node", "server.js"]