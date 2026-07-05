import express from 'express';
import { createProxyMiddleware } from 'http-proxy-middleware';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const PORT = process.env.PORT || 23200;
const API_TARGET = process.env.API_URL || 'http://127.0.0.1:23100';

const app = express();

// Proxy API calls
app.use('/v1', createProxyMiddleware({
  target: API_TARGET,
  changeOrigin: true,
}));

// Serve static files
app.use(express.static(join(__dirname, '../public')));

// Fallback to index.html
app.get('*', (req, res) => {
  res.sendFile(join(__dirname, '../public/index.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`✦ M365 Chat  http://0.0.0.0:${PORT}`);
  console.log(`  API proxy → ${API_TARGET}/v1`);
});
