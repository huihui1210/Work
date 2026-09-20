import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// 边栏插件通过“服务地址”加载本页，固定端口方便在多维表格中注册
export default defineConfig({
  // 相对路径：部署到 GitHub Pages 任意仓库子路径下资源都能正确加载
  base: './',
  plugins: [react()],
  server: {
    port: 9000,
    strictPort: true,
  },
});
