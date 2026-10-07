const fs = require('fs');
const path = require('path');

function loadEnv() {
  const env = {};
  const envFile = path.join(__dirname, '.env');
  if (fs.existsSync(envFile)) {
    fs.readFileSync(envFile, 'utf-8')
      .split(/\r?\n/)
      .forEach((raw) => {
        const line = raw.replace(/^\uFEFF/, '');
        const m = line.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)\s*$/);
        if (m) env[m[1]] = m[2].replace(/^["']|["']$/g, '');
      });
  }
  return env;
}

const env = loadEnv();

module.exports = {
  apps: [
    {
      name: 'gotongroyong',
      script: './server.js',
      cwd: __dirname,
      instances: 1,
      exec_mode: 'fork',
      autorestart: true,
      max_memory_restart: '300M',
      env: {
        NODE_ENV: 'production',
        PORT: env.PORT || '3000',
        ADMIN_PASSWORD: env.ADMIN_PASSWORD || 'admin123'
      }
    }
  ]
};