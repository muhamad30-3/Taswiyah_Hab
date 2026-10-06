module.exports = {
  apps: [{
    name: "taswiyah-api",
    script: "./artifacts/api-server/dist/index.mjs",
    node_args: "--enable-source-maps",
    instances: 1,
    exec_mode: "fork",
    autorestart: true,
    min_uptime: "10s",
    max_restarts: 10,
    restart_delay: 3000,
    exp_backoff_restart_delay: 100,
    kill_timeout: 10000,
    time: true,
    env_production: {
      NODE_ENV: "production",
    },
  }],
};
