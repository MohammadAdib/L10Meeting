import { defineConfig } from 'vite';

export default defineConfig({
  base: process.env.BUILD_TARGET === 'desktop' ? './' : '/CompanyTools/',
  build: {
    outDir: 'dist',
  },
  server: {
    proxy: {
      '/api': 'http://localhost:3847',
    },
  },
});
