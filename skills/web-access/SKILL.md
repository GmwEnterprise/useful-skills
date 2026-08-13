---
name: web-access
description: 通用网络检索技能，分多层通道指示如何进行有效的网络检索任务。适用于需要访问通用网络信息、搜索定位、抓取网页、联网核实或调研（含社交媒体、登录态站点、动态渲染页）的场景。
---

## 核心原则

四步循环：

1. **明确成功标准**：完成需要哪些字段、来源、可信度。
2. **选最可能直达的起点**：不按固定顺序遍历所有通道。
3. **过程校验**：每步结果是证据，对照标准更新判断；方向错了立即调整，不在同一方式上反复重试；搜索没命中也可能意味着目标不存在；遇到弹窗、登录墙等障碍先判断是否真挡住目标——内容可能已在 DOM 中，交互只是展示手段，挡住才处理、没挡住就绕过。
4. **完成判断**：对照标准确认完成，不过度操作。

## 信息真实性

- 一手信息优于二手：搜索引擎和聚合平台只作发现入口，定位后必须跳一手来源（官网、官方平台、原文）确认。
- 警惕循环印证：多源引用同一错误会制造真实性假象——定位到一手来源后以一手为准，不靠引用数量判断真伪。
- 禁用 AI 摘要适配器（`opencli grok/doubao/gemini` 等），其返回属于二手综述。
- 找不到官网时，权威媒体原创报道（非转载）可作次级依据，但须向用户声明转述误差。

| 信息类型 | 一手来源 |
|----------|---------|
| 政策 / 法规 | 发布机构官网 |
| 企业公告 | 公司官方新闻页 |
| 学术声明 | 原始论文 / 机构官网 |
| 工具能力 / 用法 | 官方文档 / 源码 |

## 通道分层

按目标特征选最可能直达的层：

| 层 | 触发条件 | 实现 |
|----|---------|------|
| **L1** | URL 已知的简单公开页（静态 / 服务端渲染） | 内置 WebFetch（零依赖，最快） |
| **L2** | 已知复杂站点，且 opencli 已有适配器 | `opencli <site> <command>`（结构化输出，最省 token） |
| **L3** | 不知具体目标，需搜索引擎定位 | `opencli google search`（默认）；不可用时用 `duckduckgo` / `brave` 兜底 |
| **L4** | 前述均不适用 / 失败，或必须浏览器交互、或需兜底检索 | `opencli web read`（公开页渲染）与 `opencli browser *`（通用交互），详见 opencli-browser skill |

### L1 WebFetch

- 默认入口，无需 opencli 预检。WebFetch 由小模型按 prompt 定向提取，适合从已知 URL 取特定信息，非全文 dump。
- WebFetch 不可用（网络 / 代理 / 被墙）、目标是 JS 渲染的 SPA、或需要 iframe 内嵌内容时，升级到 L4 的 `opencli web read`。

### L2 适配器要点

- **L2 永远优先于 L4**：opencli 结构化输出的网页内容远优于原始 HTML。
- **站点内交互链接可靠，手动拼 URL 不可靠**：适配器返回的 `url` 字段携带平台所需完整上下文（如小红书 `xsec_token`），优先用返回的 `url` 作为下一步入口，不自行拼 URL。
- 常见登录态社交平台（小红书、微博、知乎、B站、贴吧、虎扑、抖音、X / Twitter、Instagram、Reddit、TikTok 等）大多已有适配器，使用前先 `opencli <site> -h` 确认该站适配器是否存在。

### L3 搜索引擎

- 默认 `opencli google search`；不可用 / 被墙 / 结果不足时回退 `duckduckgo` / `brave`（`opencli -h` 顶层帮助的 Site adapters 列表查可用引擎名）。
- 搜索结果只作发现入口，定位后回 L1 或 L2 取一手内容。

### L4 浏览器执行

本层工具：`opencli web read`（公开页 Chrome 渲染，L1 升级回退）与 `opencli browser *`（通用交互）。`opencli browser *` 的命令用法（open / click / type / extract / network / screenshot / scroll、选择器契约、stale-ref、网络捕获）**加载 `opencli-browser` skill**。

**风险须知**：首次进入 L4 前，向用户简短声明「部分站点对浏览器自动化检测严格，存在账号风险，继续即视为接受」。

执行约束：

- 媒体：图片用 `extract` 取 URL 后定向读取；视频用 `/eval` 操控 `<video>`（取时长、seek 关键点、播放 / 暂停）配合 `/screenshot` 采帧分析。
- 环境：用完自建 tab 即时关闭，不操作用户已有 tab。
- 时机：部分站点对自动化检测严格，L4 仅在 L2 无适配器或拿不全字段、L3 无果、L1 不可用时使用。

DOM / 反爬事实：

- 页面常含大量已加载未展示内容（轮播非当前帧、折叠区、懒加载占位），以数据结构为单位可直接触达。
- DOM 有选择器不可跨越的边界（Shadow DOM `shadowRoot`、iframe `contentDocument`），需递归遍历。
- `/scroll` 到底触发懒加载。
- 平台「内容不存在」提示**不一定反映真实状态**，可能是 URL 缺参或触发反爬。
- 短时间密集打开页面触发反爬风控 → 控制节奏。

## 程序化 vs GUI 交互

- **程序化**（L1-L3）：快、精确、省 token，但易触发反爬。
- **GUI 交互**（L4）：慢，但最像人、确定性最高。
- GUI 是程序化的有效探测：通过真实交互观察站点行为（URL 模式、必需参数、跳转逻辑）为程序化提供依据；程序化受阻时 GUI 兜底。

## 登录判断

核心问题只有一个：**目标内容拿到了吗？**

1. 先用 L1 / L2 / L3 取内容。
2. 确认拿不到、且判断「登录能解决」时，才请用户登录：「当前页面在未登录状态下无法获取 [具体内容]，请在你的浏览器中登录 [网站名]，完成后告诉我继续。」
3. COOKIE / INTERCEPT / UI 策略适配器复用浏览器登录态，无需额外处理。
4. 登录完成后无需重启，刷新继续。

不要一上来就要求登录。

## 并行分治

存在多个独立目标时，派子 Agent 并行处理，主 Agent 只收摘要。

分治条件（同时满足才分）：

- 目标彼此独立。
- 量大 / 长任务。
- 简单单页不分。

子 Agent prompt 规范：

- **目标导向措辞**：用「获取 / 了解 / 调研 / 找到」，避免「搜索」（会锚定 L3 搜索引擎，反爬站点可能需 L2 / L4 直达）。
- **必须写明**：加载 web-access skill 并遵循其指引。
- 主 Agent 只收摘要，脏数据 / 原始抓取不进主上下文。

并行约束（opencli 非无竞态）：

- **反爬维度（同站串行）**：同一站点多次请求必须串行 + 退避，否则触发 429 / 封号；不同站点可并行。与 strategy 无关——`browser:false` 的 PUBLIC 请求同站仍要串行。
- **资源维度（Chrome 占用）**：全局一个 daemon + 一个 Chrome，所有 `browser:true` 适配器和 L4 命令经它；`browser:false`（PUBLIC）走 Node fetch 不占 Chrome。L4 并行需给每个子 Agent 不同 session 名（如 `probe-a`、`probe-b`），否则抢占同一 tab lease。

节奏：N 个子 Agent 跨 N 个不同站点并行收益最大；同站多查留在单子 Agent 内串行。

## 频率预算与失败回退

### 单题预算

「单个用户问题」= 同一意图链路的一次求解；追问 / 澄清若核心问题未变，仍算同一题。

先建台账，每次执行搜索命令后立即更新：`site` / `query` / `count` / `status`。

计数规则：

- `opencli -h`、`opencli <site> -h`、`<command> -h` 属预检，**不计**。
- 一次真正的 `opencli <site> ...` 执行 = 该站 1 次。
- 因报错 / 超时 / 验证码 / 反爬 / 登录态异常失败也算 1 次，不无限重试。

频率上限：

- 每站默认最多 **2 次**；第 2 次必须有明确理由（加时间 / 地区 / 类别 / 关键词限定）。
- 不进行第 3 次；信息不足时停止扩搜并明确说明缺口。

### 失败回退

- 单源失败不中止整体搜索。
- 回退同类其他站点，或回退 L3 搜索引擎。
- 始终以 `opencli <site> -h` 实际结果为准。
- 不假设任何站点「绝对可用」。

## 站点经验积累

按主域名存经验到 `refs/site-patterns/{domain}.md`，格式见 `refs/site-patterns/_template.md`。

**操作前必读**：确定目标站点后，若 `refs/site-patterns/{domain}.md` 存在，必须先读其平台特征 / 有效模式 / 已知陷阱作为先验；按经验失败则回退通用模式并更新文件。

域名归并：`{domain}` 用主域名（apex / eTLD+1，不含子域前缀）。opencli 命令的 `domain` 字段常含子域（如 `kyfw.12306.cn`），一律归并到主域名文件（`12306.cn.md`），子域记入 frontmatter `aliases`，不单建。

记录字段：`domain` / `aliases` / `updated` frontmatter + 平台特征 / 适配器命令（含已知坑）/ 有效模式 / 已知陷阱（含发现日期）。

原则：

- 当作「可能有效的提示」而非「保证正确」。
- 按经验失败则回退通用模式并更新文件。
- 操作成功后主动写入验证过的事实，不写未确认猜测。

## 查询结束汇报

每次查询末尾追加：

```md
搜索摘要
- 网站：<site1> | 查询词：<term1> | 次数：<n>
- 网站：<site2> | 查询词：<term2>；<term3> | 次数：<n>
- 已跳过：<site3>，原因：不可用 / 达到频率上限
```

## 范围说明

聚焦读类调研（搜索 / 抓取 / 核实 / 调研）。发布、评论、点赞等写操作不在范围；登录仅用于读取登录态内容。

opencli 无「本地书签 / 历史检索」能力（原 `find-url` 已弃）。需定位「之前看过的页面」「公司内部系统」等公网搜不到的目标时，需另行配备本地检索工具。

## opencli 基本操作

### 命令语法与 strategy

统一入口 `opencli <site> <command> [args] [options]`。每个适配器命令带 `strategy` 标签，决定是否需要浏览器：

| strategy | 需 Chrome | 含义 |
|---|---|---|
| `PUBLIC` | 否 | 纯 HTTP，无需登录 |
| `LOCAL` | 否 | 本地 / 开发端点 |
| `COOKIE` | 是 | 复用浏览器登录态 cookie |
| `INTERCEPT` | 是 | 自动开窗捕获带签名请求 |
| `UI` | 是 | 完整 DOM 交互 |

`COOKIE` / `INTERCEPT` / `UI` 需 Chrome 已登录目标站 + 已装 opencli 扩展；命令复用当前会话凭证，无需重新登录。

### 通用 flag

| flag | 作用 |
|---|---|
| `-f, --format <fmt>` | `table`（TTY 默认）/ `yaml`（非 TTY 默认）/ `json` / `plain` / `md` / `csv`。Agent 优先 `yaml` 或 `json` 拿结构化输出 |
| `-v, --verbose` | 调试日志 + 失败堆栈 |

命令专属 flag（`--limit`、`--tab`、`--filter` 等）非通用，以 `<site> <command> -h` 为准。

### 发现与预检

每次使用 opencli 前必须先做（不计入搜索次数）：

- `opencli -h` —— 顶层帮助，紧凑列出全部适配器名（Site / App / External adapters），用于定位候选站点的确切适配器名；**不要用 `opencli list`**，其全量输出超长（百万字符级）。
- 选定站点后：`opencli <site> -h` —— 查看子命令；**适配器不存在时回退为顶层帮助**，以输出首行是否为 `Usage: opencli <site> <command>` 判断存在与否。
- 需结构化细节（`browser` / `domain` / `access` / 各命令参数）：`opencli <site> --help -f yaml`。
- 锁定子命令后：`opencli <site> <command> -h` —— 查看参数、输出列、策略。

例外：L1 WebFetch 的公开页无需 opencli 预检。

禁止在文档里硬编码命令签名或假设参数；一律以 `-h` 实时输出为准，避免文档漂移。

### 排障

`COOKIE` / `INTERCEPT` / `UI` 适配器或 `opencli browser *` 失败时，先跑 `opencli doctor` 诊断浏览器桥（daemon + 扩展 + Chrome 连接）。`PUBLIC` / `LOCAL` 适配器和 `opencli -h` / `opencli <site> -h` 预检不依赖浏览器桥，无需 doctor 绿。

## 参考文件与协作 skill

- `refs/site-patterns/_template.md` —— 站点经验记录模板；站点经验按域名存于 `refs/site-patterns/{domain}.md`，按需读写。
- **`opencli-browser` skill** —— L4 浏览器执行的命令参考（opencli browser 子命令、选择器契约、网络捕获等），进入 L4 时加载。
