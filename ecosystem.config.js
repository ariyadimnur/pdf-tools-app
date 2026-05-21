module.exports = {
  apps: [{
    name: "pdf-tools-service",
    script: "./server.js",
    instances: "max",
    exec_mode: "cluster",
    watch: true,
    ignore_watch: ["uploads", "node_modules", "public"],
    max_memory_restart: "1G",
    env: {
      NODE_ENV: "production",
      PORT: 8989 // Sudah disesuaikan ke port baru
    }
  }]
};