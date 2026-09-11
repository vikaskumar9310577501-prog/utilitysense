import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

function localApiPlugin() {
  return {
    name: 'local-api-endpoints',
    configureServer(server) {
      server.middlewares.use('/api/send-otp', async (req, res) => {
        if (req.method === 'OPTIONS') {
          res.writeHead(200, {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type',
          });
          res.end();
          return;
        }

        let body = '';
        req.on('data', chunk => { body += chunk; });
        req.on('end', async () => {
          try {
            req.body = body ? JSON.parse(body) : {};
          } catch (e) {
            req.body = {};
          }

          res.status = (statusCode) => {
            res.statusCode = statusCode;
            return res;
          };
          res.json = (data) => {
            res.setHeader('Content-Type', 'application/json');
            res.end(JSON.stringify(data));
            return res;
          };

          try {
            const { default: handler } = await import('./api/send-otp.js');
            await handler(req, res);
          } catch (err) {
            console.error('Local OTP API error:', err);
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: err.message || 'Failed to send OTP' }));
          }
        });
      });
    }
  };
}

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react(), localApiPlugin()],
})
