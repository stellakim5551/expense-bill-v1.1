/** @type {import('next').NextConfig} */
const nextConfig = {
  output: 'export', // 关键：启用静态导出
  basePath: '/expense-bill-v1.1', // 你的仓库名称（必须与 GitHub 仓库名一致）
}

export default nextConfig