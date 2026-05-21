import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import {defineConfig, loadEnv} from 'vite';

export default defineConfig(({mode}) => {
  const env = loadEnv(mode, '.', '');
  return {
    plugins: [react(), tailwindcss()],
    define: {
      'process.env.GEMINI_API_KEY': JSON.stringify(process.env.GEMINI_API_KEY || env.GEMINI_API_KEY),
      'process.env.APP_URL': JSON.stringify(process.env.APP_URL || env.APP_URL),
      'process.env.POWER_AUTOMATE_URL': JSON.stringify(process.env.POWER_AUTOMATE_URL || env.POWER_AUTOMATE_URL),
      'process.env.POWER_AUTOMATE_EMAIL_URL': JSON.stringify(process.env.POWER_AUTOMATE_EMAIL_URL || env.POWER_AUTOMATE_EMAIL_URL),
      'process.env.POWER_AUTOMATE_TEAMS_URL': JSON.stringify(process.env.POWER_AUTOMATE_TEAMS_URL || env.POWER_AUTOMATE_TEAMS_URL),
      'process.env.TEAMS_SCREENSHOT_SCALE': JSON.stringify(process.env.TEAMS_SCREENSHOT_SCALE || env.TEAMS_SCREENSHOT_SCALE),
    },
    build: {
      rollupOptions: {
        input: {
          main: path.resolve(__dirname, 'index.html'),
          admin: path.resolve(__dirname, 'admin.html'),
        },
      },
    },
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modifyâfile watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
      // Disable file watching when DISABLE_HMR is true to save CPU during agent edits.
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
  };
});
