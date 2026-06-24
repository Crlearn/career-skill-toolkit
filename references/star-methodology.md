# STAR 重写规则

## STAR 篇幅分配

- **S（情境）**：1 行，项目背景和挑战，不铺陈背景故事
- **T（任务）**：1 行，角色和负责范围
- **A（行动）**：2-3 行，具体做了什么（**核心**，占 50%+ 篇幅）
- **R（结果）**：1 行，量化效果，缺数据标注 `[待补充]`

---

## 替代框架

### X-Y-Z 公式（Google 方法）

结构："Accomplished [X] as measured by [Y] by doing [Z]"
- X = 取得了什么成果
- Y = 如何衡量（数据）
- Z = 采取了什么行动

```
❌ BEFORE: "Managed social media accounts"
✅ AFTER: "Grew Instagram following by 250% (5K to 17.5K) by implementing daily content calendar and influencer partnerships"
```

### CAR 方法（Challenge-Action-Result）

- **C**hallenge：存在什么问题
- **A**ction：你做了什么
- **R**esult：发生了什么

```
"Reduced customer churn (C) by implementing proactive outreach program (A), retaining 85% of at-risk accounts worth $500K ARR (R)"
```

### STAR+R 方法（进阶版，适用于资深岗位）

在标准 STAR 基础上增加 **Reflection（反思）** 维度：
- **R**eflection：从这段经历中学到了什么？如果重来会怎么做？

Reflection 维度传递资历信号：初级候选人描述发生了什么，资深候选人提炼经验教训。

---

## 重写约束

- 只改表达方式，不改事实内容
- 不编造不存在的经历、数据或成果
- 缺少量化数据时保留原有表述，标注 `[待补充]`
- 关键词要自然融入，不能生硬堆砌（同一关键词全文最多出现 2-3 次）
- 每段经历控制在 3-5 行
- 总经历段数 **最多 3 段**，选择与目标岗位最匹配的

---

## 关键词融入策略

### 融入方式
- **自然嵌入（推荐）**：将关键词放入 Action 描述中，使其成为技术选型或方法论的一部分
- **技能清单覆盖**：确保"岗位技能"模块包含所有 10 个关键词
- **避免堆砌**：不在不适合的语境中强行插入，关键词应服务于内容表达

### 关键词密度指南
- 关键关键词：全文出现 2-4 次
- 重要关键词：全文出现 1-2 次
- 变换表述（如 "led team" 和 "team leadership"），避免重复感

### 合法改写原则

只改词汇不改事实。将用户真实经历用 JD 的精确术语重新表述：

| JD 关键词 | 简历原文 | 合法改写 |
|-----------|---------|---------|
| "RAG pipelines" | "LLM workflows with retrieval" | → "RAG pipeline design and LLM orchestration workflows" |
| "MLOps" | "observability, evals, error handling" | → "MLOps and observability: evals, error handling, cost monitoring" |
| "stakeholder management" | "collaborated with team" | → "stakeholder management across engineering, operations, and business" |

**核心原则：绝不添加候选人没有的技能。只用 JD 词汇重新表述真实经历。**

### "6 秒招聘者扫描"关键词分布策略
- **Professional Summary**：植入前 5 个最重要的关键词
- **每段经历的首个 bullet**：植入 1-2 个关键词（招聘者通常只读每段第一行）
- **Skills 模块**：覆盖剩余关键词

---

## 重构示例

```
【原文】负责公司官网改版项目
【STAR重构】主导公司官网全面改版（S/T），采用 React + Node.js 技术栈重构前端架构（A），
页面加载速度提升60%，用户转化率提高25%（R）
```

---

## 自我评价重写

如果用户原始简历中包含自我评价内容，则按以下规则重写：

**重写原则**：
1. **不超过 3 行**，控制在 60-100 字
2. 内容围绕三个维度组织，每维度一句话：
   - **我有什么**：与岗位直接相关的经历或资源
   - **我擅长什么**：差异化优势或突出能力
   - **为什么适合**：呼应 JD 的核心需求或隐藏偏好
3. 不写空泛的「性格开朗」「吃苦耐劳」等无效描述
4. 优先使用 JD 中的关键词，自然融入而非堆砌

**示例**（短剧内容岗）：
> 热爱短剧与互动影游内容，日常高频关注抖音、小红书等平台的短剧赛道动态与爆款趋势；熟练使用即梦、豆包等 AI 工具辅助内容创作，乐于探索新型内容生产方式；执行力强，能快速适应节奏并接受反馈，注重细节与版本管理。

如果用户原始简历没有自我评价，且 JD 分析中存在明显的隐藏偏好可以通过自我评价补充，则主动建议添加；否则不强制生成。
