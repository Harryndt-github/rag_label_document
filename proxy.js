// Simple CORS proxy for LM Studio
// Forwards requests from the browser to LM Studio at localhost:1234
const http = require('http');

const LM_STUDIO = { host: '127.0.0.1', port: 1234 };
const PROXY_PORT = 1235;

const server = http.createServer((req, res) => {
    // CORS headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }

    let body = [];
    req.on('data', chunk => body.push(chunk));
    req.on('end', () => {
        const bodyBuffer = Buffer.concat(body);
        const options = {
            hostname: LM_STUDIO.host,
            port: LM_STUDIO.port,
            path: req.url,
            method: req.method,
            headers: { ...req.headers, host: `${LM_STUDIO.host}:${LM_STUDIO.port}` }
        };

        const proxy = http.request(options, proxyRes => {
            res.writeHead(proxyRes.statusCode, {
                ...proxyRes.headers,
                'Access-Control-Allow-Origin': '*'
            });
            proxyRes.pipe(res);
        });

        proxy.on('error', err => {
            res.writeHead(502);
            res.end(JSON.stringify({ error: 'LM Studio not available', detail: err.message }));
        });

        if (bodyBuffer.length > 0) proxy.write(bodyBuffer);
        proxy.end();
    });
});

server.listen(PROXY_PORT, () => {
    console.log(`CORS Proxy running on http://localhost:${PROXY_PORT} → LM Studio at ${LM_STUDIO.host}:${LM_STUDIO.port}`);
});
