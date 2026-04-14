// ═══════════════════════════════════════════════════════════
//  Unified CORS Proxy for LM Studio + Docling Server
//  Forwards:
//    /api/docling/*  → Docling Server (localhost:5050)
//    /v1/*           → LM Studio (localhost:1234)
// ═══════════════════════════════════════════════════════════
const http = require('http');
const FormData = require('form-data');

const LM_STUDIO = { host: '127.0.0.1', port: 1234 };
const DOCLING = { host: '127.0.0.1', port: 5050 };
const PROXY_PORT = 1235;

const server = http.createServer((req, res) => {
    // CORS headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, PUT, DELETE');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }

    // ── Route: Docling endpoints → port 5050 ──
    const isDocling = req.url.startsWith('/api/docling');

    let body = [];
    req.on('data', chunk => body.push(chunk));
    req.on('end', () => {
        const bodyBuffer = Buffer.concat(body);

        const target = isDocling ? DOCLING : LM_STUDIO;
        const options = {
            hostname: target.host,
            port: target.port,
            path: req.url,
            method: req.method,
            headers: { ...req.headers, host: `${target.host}:${target.port}` }
        };

        const proxy = http.request(options, proxyRes => {
            // Increase max listener limit for large responses
            proxyRes.setMaxListeners(20);

            res.writeHead(proxyRes.statusCode, {
                ...proxyRes.headers,
                'Access-Control-Allow-Origin': '*'
            });
            proxyRes.pipe(res);
        });

        proxy.on('error', err => {
            const serviceName = isDocling ? 'Docling Server' : 'LM Studio';
            res.writeHead(502);
            res.end(JSON.stringify({
                error: `${serviceName} not available`,
                detail: err.message,
                service: isDocling ? 'docling' : 'lm-studio'
            }));
        });

        if (bodyBuffer.length > 0) proxy.write(bodyBuffer);
        proxy.end();
    });
});

server.listen(PROXY_PORT, () => {
    console.log(`═══════════════════════════════════════════════════════`);
    console.log(`  Unified CORS Proxy — Port ${PROXY_PORT}`);
    console.log(`  /api/docling/*  → Docling Server (${DOCLING.host}:${DOCLING.port})`);
    console.log(`  /v1/*           → LM Studio (${LM_STUDIO.host}:${LM_STUDIO.port})`);
    console.log(`═══════════════════════════════════════════════════════`);
});
