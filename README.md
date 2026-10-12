# 美元交易明细人民币换算（USD → CNY）

纯前端小工具：上传 **CSV / XLSX** 交易明细，按**数据日期当天**的美元兑人民币汇率，
自动为每一行补出 **「汇率」** 和 **「总计 (RMB)」** 两列，并给出合计与导出文件。

汇率来源：[frankfurter.dev](https://frankfurter.dev/)（开源汇率 API，无需密钥，支持 CORS，浏览器可直接调用）。

[![Deploy to Cloudflare](https://deploy.workers.cloudflare.com/button)](https://deploy.workers.cloudflare.com/?url=https://github.com/CarlJia/usd-cny-fx-converter)

> **在线体验**：<https://usd-cny-fx-converter.carl-jia.workers.dev>
> 点上面的按钮，几秒钟就能把它部署到你自己的 Cloudflare 账号。

---

## 特性

- **纯静态、纯前端**：文件解析与计算全部在浏览器本地完成，明细数据不会上传到任何服务器。
- **多文件 / 多工作表**：一次可上传多个文件；XLSX 的每个工作表都会被解析，可逐个勾选是否参与换算。
- **字段自动识别**：自动定位「日期列」与「美元金额列」，识别不准可在界面上手动改。
- **日期格式兼容**：Excel 序列号（含 1904 日期系统）、`2026-07-31`、`2026/7/31`、`2026年7月31日`、`7/31/2026`、`20260731`、ISO 时间串等。
- **逐日取汇率**：按每一行的日期取当天汇率；当天无牌价（周末 / 节假日）时沿用最近一个工作日的汇率，并在「汇率取值日」列标出真实取值日期，表格里以黄色底纹提示。
- **本地缓存**：查过的汇率存进 localStorage，重复换算不重复请求。
- **导出**：一键导出 XLSX（含 `汇总` + `全部明细` + 按原工作表拆分）或 CSV（带 BOM，Excel 直接打开不乱码）。
- **命令行版本**：`scripts/convert-cli.cjs`，同一套核心逻辑，适合批量 / 定时任务。

## 计算公式

```
汇率      = frankfurter USD→CNY 在「数据日期」当日的参考价
总计 (RMB) = 总计 (USD) × 汇率   （四舍五入到设定的小数位，默认 2 位）
```

---

## 目录结构

```
fx-converter/
├── public/                   # 部署目录（Cloudflare 静态资源）
│   ├── index.html
│   ├── styles.css
│   ├── core.js               # 核心逻辑：日期解析、列识别、汇率获取与匹配（无 DOM 依赖）
│   ├── app.js                # 页面交互与编排
│   └── vendor/
│       └── xlsx.full.min.js  # SheetJS 0.20.3（读写 XLSX，Apache-2.0）
├── scripts/
│   ├── convert-cli.cjs       # 命令行换算
│   └── serve.cjs             # 零依赖本地静态服务器
├── test/
│   ├── core.test.cjs         # 单元测试（node --test）
│   └── verify-output.cjs     # 读回导出的 XLSX 做校验
├── .github/workflows/
│   └── deploy.yml            # GitHub Actions：push 到 main 自动部署
├── wrangler.jsonc            # Cloudflare 部署配置
└── package.json
```

---

## 本地运行

```bash
# 方式一：零依赖静态服务器
node scripts/serve.cjs            # http://localhost:5173

# 方式二：wrangler 本地预览（环境与线上一致）
npx wrangler dev
```

打开页面 → 拖入 `.xlsx` / `.csv` → 核对字段 → 开始换算 → 导出。

## 命令行用法

```bash
node scripts/convert-cli.cjs "7-9月交易明晰.xlsx" \
  -o "7-9月交易明晰-CNY换算.xlsx" \
  --base USD --quote CNY \
  --rate-decimals 4 --amount-decimals 2

# 只看结果不写文件
node scripts/convert-cli.cjs data.xlsx --json
```

可选参数：`--date-col <列名>`、`--amount-col <列名>` 用于在自动识别不准时手动指定列。

## 测试

```bash
node --test test/core.test.cjs
node test/verify-output.cjs "outputs/7-9月交易明晰-CNY换算.xlsx"
```

---

## 部署到 Cloudflare

### 方案 A：一键部署（推荐，点按钮就行）

[![Deploy to Cloudflare](https://deploy.workers.cloudflare.com/button)](https://deploy.workers.cloudflare.com/?url=https://github.com/CarlJia/usd-cny-fx-converter)

点击按钮后 Cloudflare 会自动完成：

1. **克隆仓库** —— 把本项目复制一份到你自己的 GitHub 账号，之后你可以在自己的仓库里继续改；
2. **确认配置** —— 只需要确认仓库名 / Worker 名，其余留默认即可；
3. **构建并部署** —— 通过 Workers Builds 执行 `npm run deploy`（即 `npx wrangler deploy`），部署到 `https://<Worker 名>.<你的子域>.workers.dev`；
4. **开通 CI/CD** —— 之后每次 push 到 `main` 会自动重新部署，向其他分支推送还会生成预览地址。

本项目**没有任何 KV / D1 / R2 绑定，也没有环境变量或密钥**，所以不需要额外配置，一路确认即可完成。
默认构建命令留空、部署命令 `npx wrangler deploy`，`wrangler.jsonc` 里的 `assets.directory` 已指向 `./public`。

> 前提：Deploy to Cloudflare 按钮只支持 **Workers 应用**（本项目正是 Worker + 静态资源，不是 Pages 项目），
> 且仓库必须是**公开**的 Git 仓库（github.com / gitlab.com）。

### 方案 B：GitHub Actions 自动部署（push 即发布）

仓库里已经放好工作流 `.github/workflows/deploy.yml` —— **push 到 `main` 就自动跑一次 `wrangler deploy`**，
把最新代码发布到你自己账号下的 Worker。也可以在 Actions 页面手动触发。

| 触发事件 | 行为 |
| --- | --- |
| push 到 `main` | 先跑单元测试，通过后执行 `wrangler deploy` 发布到线上 |
| 向 `main` 提 PR | 只上传预览版本（`wrangler preview`），**不会动线上**，预览地址写在运行摘要里 |
| 手动触发 | Actions 页面 → 选中工作流 → *Run workflow* |

工作流内置了 `concurrency`，同一分支上新的推送会取消上一次未跑完的部署；测试不通过则不会发布。

**首次使用需要配置 2 个仓库密钥**（Settings → Secrets and variables → Actions → New repository secret）：

| Secret | 值 |
| --- | --- |
| `CLOUDFLARE_API_TOKEN` | 具备 **Workers Scripts:Edit** 权限的 API Token |
| `CLOUDFLARE_ACCOUNT_ID` | 你的 Cloudflare Account ID（`npx wrangler whoami` 可查） |

创建 API Token 的最短路径 —— 下面这个链接已经把**权限和账号预填好**，点开只需 *Continue to summary* → *Create Token*：

<https://dash.cloudflare.com/profile/api-tokens?permissionGroupKeys=%5B%7B%22key%22%3A%22workers_scripts%22%2C%22type%22%3A%22edit%22%7D%5D&accountId=e4a57185c8f9ea8de6a1c8f4cd1c5681&name=github-actions-usd-cny-fx-converter>

也可以直接用 Cloudflare 面板里的模板「**Edit Cloudflare Workers**」。

写入密钥（用 `gh` 命令行，输入的值不会回显，也不会进 shell 历史）：

```bash
gh secret set CLOUDFLARE_API_TOKEN -R <你的账号>/usd-cny-fx-converter
gh secret set CLOUDFLARE_ACCOUNT_ID -R <你的账号>/usd-cny-fx-converter
```

之后随便 push 一次 `main`，或到 Actions 页面手动跑一次即可。

> **和方案 A 的区别**：方案 A 走 Cloudflare 自己的 Workers Builds，构建不在 GitHub 上跑；
> 方案 B 走 GitHub Actions，日志、权限、触发条件都在你自己的 GitHub 仓库里，
> 适合想统一在 GitHub 上看 CI，或后面要加 lint / 测试 / 多环境部署的场景。两者可以只选一个。

### 方案 C：本地命令行部署

```bash
npx wrangler login          # 首次需要授权
npx wrangler deploy
```

部署完成后会得到 `https://usd-cny-fx-converter.<你的子域>.workers.dev`。

### 方案 D：Cloudflare Pages

```bash
npx wrangler pages deploy public --project-name=usd-cny-fx-converter
```

或者在 Cloudflare 控制台 → Workers & Pages → Create → Pages → 连接 Git 仓库：

| 配置项 | 值 |
| --- | --- |
| Framework preset | None |
| Build command | 留空 |
| Build output directory | `public` |

### 方案 E：拖拽上传

Cloudflare 控制台 → Workers & Pages → Create → Pages → **Upload assets**，
直接把 `public/` 目录整个拖上去即可。

> 说明：本项目是纯静态资源，不需要任何服务端逻辑，也不需要配置 API 密钥。
> 汇率请求由浏览器直接发往 `api.frankfurter.dev`（该接口响应头为 `Access-Control-Allow-Origin: *`）。

---

## 汇率接口说明

| 用途 | 请求 |
| --- | --- |
| 区间汇率（首选） | `https://api.frankfurter.dev/v2/rates?from=2026-07-01&to=2026-09-30&base=USD&quotes=CNY` |
| 回退接口（v2 不可用时） | `https://api.frankfurter.dev/v1/2026-07-01..2026-09-30?base=USD&symbols=CNY` |

响应：

```json
[{ "date": "2026-07-31", "base": "USD", "quote": "CNY", "rate": 6.7468 }]
```

- 取值口径为**当日参考价**，不是实时成交价，适合对账 / 报表 / 报税场景。
- 长区间会自动按 366 天分片请求。
- 接口无需密钥、无配额限制，但建议保留本地缓存。

---

## 已知边界

- 「总计 (RMB)」按逐行四舍五入后再求和，与「先求和再乘汇率」可能存在几分钱的差异（这是逐行换算的正常现象）。
- 汇率取的是**数据日期**当天，而非平台结算日 / 到账日；若业务口径要求按结算日换算，可把明细里的结算日列指定为「日期列」。
- CSV 编码支持 UTF-8 与 GBK 自动识别；其他编码可能需要在导出前转成 UTF-8。

## 许可

代码 MIT。汇率数据版权归各数据提供方所有，详见 [frankfurter.dev/license](https://frankfurter.dev/license/)。
