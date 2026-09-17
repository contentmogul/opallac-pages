FROM node:20-slim
WORKDIR /app
ENV NODE_ENV=production
COPY package*.json ./
RUN npm ci --only=production
COPY src ./src
EXPOSE 8080
CMD ["node", "src/index.js"]
