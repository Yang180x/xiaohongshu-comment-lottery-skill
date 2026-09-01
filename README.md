# 🎯 小红书评论区寄样用户筛选 Skill

基于小红书笔记评论区，自动读取公开评论和候选用户主页信息，筛选适合 **寄样、体验官招募、产品试用、KOC 种草** 的用户，并生成 Excel 筛选报告。

本 Skill 不采用随机抽取，也不会单纯按照粉丝数排序。

> **真实需求 > 评论相关度 > 产品匹配度 > 账号质量 > 内容价值 > 粉丝数量**

---

## ✨ 能做什么

只需要提供：

* 🔗 小红书笔记链接
* 👥 需要筛选的人数

例如：

```text
帮我从这篇笔记的评论区筛选 10 位适合寄样的用户：
https://www.xiaohongshu.com/explore/xxxx（相关小红书笔记链接）
```

Skill 会自动完成：

```text
📝 分析笔记和产品
        ↓
💬 读取评论区
        ↓
🧹 清洗无效评论
        ↓
🎯 筛选高相关用户
        ↓
👤 分析候选用户主页
        ↓
📊 综合评分与排名
        ↓
📦 正式名单 + 备选名单
        ↓
📑 生成 Excel 报告
```

适用于：

* 📦 产品寄样
* 🧪 新品试用
* 🙋 体验官招募
* 🌱 KOC / 素人种草
* 🔍 评论区潜在消费者筛选

---

## 📊 输出内容

默认生成一个包含 3 个 Sheet 的 Excel：

| Sheet     | 内容              |
| --------- | --------------- |
| ✅ 正式名单    | 最终推荐的寄样 / 体验官用户 |
| 🕒 备选名单   | 默认额外保留 5 位候选用户  |
| 📋 评论筛选明细 | 评论评分、筛选过程和淘汰原因  |

正式名单会包含：

**昵称、主页链接、原始评论、粉丝数、账号定位、评分、推荐理由等信息。**

---

## 🤖 支持的 AI Agent

本 Skill 不限定 WorkBuddy，只要 AI Agent 支持 **Shell / CLI** 并可以调用 BrowserSkill 即可使用。

支持例如：

* Cursor
* Claude Code
* Codex
* OpenClaw
* CodeBuddy
* WorkBuddy
* Pi
* Hermes Agent
* DeepSeek Harness
* 其他支持 Skill / Shell 的 AI Agent

---

# ⚙️ 前置条件

## 1️⃣ 安装 BrowserSkill

本 Skill 依赖腾讯开源项目：

👉 https://github.com/Tencent/BrowserSkill

🛠️ 推荐直接告诉你的 AI Agent：

```text
按照 https://raw.githubusercontent.com/Tencent/BrowserSkill/main/AGENT_INSTALL.md
的说明，在本机安装并配置 browser-skill。
```

也可以手动安装。

### macOS / Linux

```bash
curl -fsSL https://raw.githubusercontent.com/Tencent/BrowserSkill/main/install.sh | sh
bsk install-skill --yes
```

### Windows PowerShell

```powershell
irm https://raw.githubusercontent.com/Tencent/BrowserSkill/main/install.ps1 | iex
bsk install-skill --yes
```

安装完成后执行：

```bash
bsk doctor
```

检查 BrowserSkill 是否配置正常。

---

## 2️⃣ 安装 BrowserSkill 浏览器插件

还需要在 **Chrome 或 Edge** 中安装 BrowserSkill 浏览器扩展。

### 🌐 Chrome / Chromium

https://chromewebstore.google.com/detail/hhcmgoofomhgciiibhipgmgkgnoenaoi

### 🌐 Microsoft Edge

https://microsoftedge.microsoft.com/addons/detail/browserskill/emacgiaaaiojkkpkddmmdfhmokgmnikg

安装完成后打开 BrowserSkill 扩展，确认连接正常，再执行：

```bash
bsk doctor
```
<img width="2560" height="1410" alt="a1f2c4a8-f3fa-4acb-98f5-ddcfb5dab154" src="https://github.com/user-attachments/assets/a85292fb-b90d-4eb5-8670-8a3727d7435b" />


---

## 3️⃣ 登录小红书

在安装 BrowserSkill 扩展的 Chrome / Edge 浏览器中登录：

```text
https://www.xiaohongshu.com/
```

Skill 会通过 BrowserSkill 使用你当前 **已经登录的真实浏览器环境** 读取页面。

---

## 4️⃣ Python 环境

生成 Excel 报告需要：

```text
Python 3.10+
openpyxl
```

安装：

```bash
pip install openpyxl
```

---

# 📦 安装 Skill

克隆仓库：

```bash
git clone https://github.com/Yang180x/xiaohongshu-comment-lottery-skill.git
```

然后将仓库放入对应 AI Agent 的 Skill 目录。

例如 WorkBuddy：

```bash
git clone https://github.com/Yang180x/xiaohongshu-comment-lottery-skill.git \
~/.workbuddy/skills/小红书评论区寄样用户筛选
```

其他 AI Agent 按照自身 Skill / Rules / Agent 配置方式安装即可。

---

# 🚀 使用方式

安装完成后，直接告诉 AI Agent：

```text
帮我从这篇小红书笔记评论区筛选 10 位适合寄样的用户：

https://www.xiaohongshu.com/explore/xxxxxxxx
```

也可以增加具体要求：

```text
帮我筛选 10 位寄样用户。

要求：
- 优先真实有产品需求的用户
- 排除纯抽奖评论和营销号
- 优先内容与产品相关的账号
- 粉丝数只作为辅助参考
- 正式名单 10 位
- 额外准备 5 位备选
```

AI Agent 会自动调用 BrowserSkill 完成评论读取、用户分析、筛选和 Excel 报告生成。

---

## 📁 仓库结构

```text
.
├── SKILL.md
├── references/
│   └── 评分与筛选细则.md
├── scripts/
│   └── generate_excel.py
└── README.md
```

---

## 📴 离线模式

如果暂时无法使用 BrowserSkill，也可以直接提供已有的评论 Excel。

Skill 可以继续完成：

```text
评论清洗 → 评论评分 → 候选排序 → Excel 报告生成
```

由于无法实时访问用户主页，账号质量、内容价值、粉丝数等数据可能不完整。

---

## ⚠️ 注意

本 Skill 用于辅助整理和分析小红书公开评论及公开主页信息。

请遵守平台规则以及相关隐私和数据保护规定，正式寄样前建议进行人工复核。

---

## 📄 License

MIT
