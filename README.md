# 美元交易明细人民币换算（USD → CNY）

纯前端小工具：上传 **CSV / XLSX** 交易明细，按**数据日期当天**的美元兑人民币汇率，
自动为每一行补出 **「汇率」** 和 **「总计 (RMB)」** 两列，并给出合计与导出文件。

汇率来源：[frankfurter.dev](https://frankfurter.dev/)（开源汇率 API，无需密钥，支持 CORS，浏览器可直接调用）。

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

### 方案 A：Workers 静态资源（推荐，本仓库已配好）

```bash
npx wrangler login          # 首次需要授权
npx wrangler deploy
```

部署完成后会得到 `https://usd-cny-fx-converter.<你的子域>.workers.dev`。
`wrangler.jsonc` 里的 `assets.directory` 已指向 `./public`，无需构建步骤。

### 方案 B：Cloudflare Pages

```bash
npx wrangler pages deploy public --project-name=usd-cny-fx-converter
```

或者在 Cloudflare 控制台 → Workers & Pages → Create → Pages → 连接 Git 仓库：

| 配置项 | 值 |
| --- | --- |
| Framework preset | None |
| Build command | 留空 |
| Build output directory | `public` |

### 方案 C：拖拽上传

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
