# 小红书评论区寄样用户筛选 Skill

基于小红书笔记评论区，自动读取公开评论与用户主页数据，按真实需求、评论相关度等六维优先级筛选寄样/体验官人选，输出正式名单+备选+明细的 Excel 报告。

## 这是什么

一个 [WorkBuddy](https://www.workbuddy.cn) / AI Agent Skill：输入一篇小红书笔记链接和目标人数，通过腾讯 BrowserSkill（bsk）操作已登录的真实浏览器，完整读取目标笔记与公开评论区，自动筛选出适合**寄样 / 体验官招募 / 试用活动**的用户。

**核心目标不是抽人，也不是按粉丝数排序。** 综合优先级恒定为：

> 真实需求 > 评论相关度 > 产品匹配度 > 账号质量 > 内容价值 > 粉丝数量

粉丝数仅作辅助因素，绝不作为主要筛选依据。

## 工作流程（7 个阶段）

```
用户输入链接 + 人数
→ 阶段1 笔记分析（产品画像）
→ 阶段2 完整读取评论区
→ 阶段3 无效评论清洗 + 评论评分
→ 阶段4 建立预筛候选池（约 2 倍）
→ 阶段5 主页分析 + 二次筛选（动态递补）
→ 阶段6 综合排名 + 正式/备选名单
→ 阶段7 生成 Excel 交付
```

## 输出

一份包含 3 个 Sheet 的 Excel 报告：

| Sheet | 内容 |
|-------|------|
| 正式名单 | 最终确定的寄样/体验官用户 |
| 备选名单 | 默认 5 位备选用户 |
| 评论筛选明细 | 全部候选评论的评分与筛选过程 |

## 仓库结构

```
.
├── SKILL.md                      # Skill 主文档（含完整执行流程与规范）
├── references/
│   └── 评分与筛选细则.md          # 六维评分标准与权重细则
└── scripts/
    └── generate_excel.py         # Excel 报告生成脚本
```

## 前置条件

- [WorkBuddy](https://www.workbuddy.cn)（或其他支持 Skill 的 AI Agent 平台）
- 腾讯 BrowserSkill（bsk CLI）+ 已登录小红书的浏览器（Chrome / Edge 安装 BrowserSkill 扩展）
- Python 3.10+ 与 `openpyxl`

## 使用方法

1. 将本仓库克隆到本地 Skill 目录：

   ```bash
   git clone https://github.com/Yang180x/xiaohongshu-comment-lottery-skill.git \
     ~/.workbuddy/skills/小红书评论区寄样用户筛选
   ```

2. 在 WorkBuddy 对话中 @skill 调用，或直接说：「帮我从这篇笔记的评论区筛选 10 位寄样用户：<笔记链接>」

## 离线模式

若 bsk 不可用，可由用户提供评论 Excel 走离线评分模式（评分逻辑复用，跳过实时读取）。

## License

MIT
