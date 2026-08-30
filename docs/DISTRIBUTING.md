# 分发构建指南（二次开发）

把 GenOffice 变成你自己的可下载安装包，需要三步：**上传代码 → 本机构建 → 发布安装包**。
本文假设你在 Windows 本地构建（已装 Node ≥ 22、Rust 工具链 + MSVC、`zip`）。

## 一、上传到你的 GitHub

1. 在 github.com 新建一个**空仓库**（不要勾选 README / .gitignore / license）。
2. 在项目根目录执行：

   ```bash
   git init
   git add .
   git commit -m "Initial import of GenOffice"
   git branch -M main
   git remote add origin https://github.com/<你的用户名>/<仓库名>.git
   git push -u origin main
   ```

   项目已内置 `.gitignore`（含 `node_modules`、`out/`、`release/`、Rust `target/` 等），
   不会把依赖和构建产物传上去。

## 二、构建可下载的安装包

```bash
npm install                 # 首次（Windows 下首次会下载 Electron，耐心等待）
npm run dist:win            # 生成 Windows 安装包（NSIS .exe）
```

产物在 `apps/shell/release/`：一个 `GenOffice-Setup-*.exe` 安装包。
`dist:win` 会自动执行：第三方许可清单 → 构建全部六个模块 → 打包。sidecar 由
`build:all` 里的 `native:build` 一并编译（首次 1-3 分钟）。

其它平台（在对应系统上执行）：

```bash
npm run dist:mac            # macOS（需 Xcode 命令行工具 + 签名证书，否则跳过公证）
npm run dist:linux          # Linux（AppImage + deb + rpm）
```

## 三、构建期注入变量（fork 后按需设置，均为可选的 HTTPS 地址）

| 环境变量 | 作用 | 不设时 |
|---|---|---|
| `GENOFFICE_UPDATE_URL` | 自动更新 feed 地址 | 禁用自动更新 |
| `GENOFFICE_DOWNLOAD_PAGE_URL` | 手动下载兜底页 | 指向官方 releases 页 |
| `GENOFFICE_GA4_MEASUREMENT_ID` / `_API_SECRET` | 匿名统计（GA4） | 统计完全关闭 |
| `GENOFFICE_FONT_CDN_URL` | 字体目录 CDN | 隐藏字体下载入口 |

设置方式：CI 里作为 secret，本地写在 `apps/shell/electron-builder.env`（已 gitignore）。

## 四、注意事项

- **未签名安装包**：个人 fork 构建的 `.exe` 没有代码签名，Windows SmartScreen 会提示
  “未知发布者”，用户点“仍要运行”即可。正式分发建议配置代码签名证书（Windows）与
  公证（macOS）。
- **不要漏跑 `native:build`**：Windows 安装包依赖 `xlsx-sidecar.exe`（Rust 编译产物，
  `dist:win` 会自动触发）；如果 sidecar 缺失，安装包会静默少带一个组件，打开表格会失败。
- **e2e 自检**（可选）：`$env:GENOFFICE_E2E_VIDEO='0'; npm run test:e2e`，应先执行
  `npm run build:all` 与各包的 `npm run fixtures`。
- 想彻底改名换标（应用名 / appId / Linux 包元数据），改动集中在
  `apps/shell/package.json`、`apps/shell/electron-builder.cjs` 两处。
