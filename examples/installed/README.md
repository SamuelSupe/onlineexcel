# 安装包生产集成示例

在本目录执行，安装已发布的 npm 包并构建示例：

```sh
npm install
npm run build
```

通过 HTTP 静态服务运行 dist/index.html，点击 Run SDK integration checks。示例使用已安装的 onlineexcel 包、复制后的 Worker、可信业务模块、React/Vue 薄适配器和严格 CSP；不引用库源码。验证尚未发布的本地改动时，先在仓库根目录构建并 `npm pack --pack-destination artifacts`，再运行下方脚本。

仓库验证可直接运行 `node scripts/build-sdk-example.mjs`，它在 `.cache/sdk-example` 创建独立安装环境。使用现有 QA 服务打开 `http://127.0.0.1:5174/.cache/sdk-example/dist/index.html`。
