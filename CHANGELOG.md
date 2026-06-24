# 更新日志

## v1.1.0 — 2026-06-25

### 修复
- 隐私脱敏：文档和模板中所有具体个人信息替换为通用占位符（XXXX-XX / 13XXXXXXXXX / xxx@example.com）
- .gitignore 改为通用规则，移除个人姓名
- 求职类型补充"兼职"选项
- `_esc()` 添加 XSS 安全注释，明确用户输入转义要求
- `_diff_resume()` 复合键改用元组 `(company, role, period)`，避免同公司不同角色冲突
- `_map_dict()` 新增大小写不敏感匹配，统一与 `_map_field()` 逻辑
- `_cmd_version()` 重构参数传递，不再直接操作 `sys.argv`
- `version restore` 默认输出到脚本目录，保持路径一致
- `validate_resume()` 统一返回 `dict` 类型，消除返回值不一致
- 提取 `_safe_filename()` 工具函数，消除重复代码
- .gitignore 收窄通配规则 `*_resume_*.json` 替代 `*_*.json`
- `read_docx()` 改为返回 dict 结构，不再静默吞掉异常
- 从 Git 历史中彻底清除敏感文件（简历 JSON/HTML、原始 PDF）

### 新增
- 字段映射系统（`map` 命令）：中英文别名自动映射为标准模板字段
- 版本管理（`version save/list/diff/restore`）：简历版本保存、对比、回滚
- `--pages N` 支持任意正整数页数（不再限制 1 或 2）
- `--versions-dir <path>` 支持自定义版本存储目录
- `version restore` 支持可选输出路径参数
- `experience[].tags` 字段：与 JD 关键词匹配的标签
- `core_advantages` 字段：核心优势模块，显示在 Header 下方

### 规则优化
- 打招呼用语改为核心优势确定后自动生成（原为可选，需用户要求）
- Step 2.5 新增内置推荐机制：AI 主动推荐核心优势组合方案，用户选择即可
- 新增流程回归规则：AI 偏离预期时必须重新阅读文档确认流程
- 新增单页内容填充规则：单页简历应尽量填满，底部空白不超过 15%
- 示例/模板中具体项目名和技术栈替换为通用占位符（XX项目/XX框架/XX工具A等）

### 文档
- 同步更新 SKILL.md、PROMPT-MERGED.md、README.md、career-tracker.md、boss-greeting.md

## v1.0.0 — 2026-06-24

### 新增
- 简历优化主 Skill（SKILL.md）：去AI化规则、JD硬技能提取、STAR重写、证据链分级、真相边界降级表
- 面试准备 Skill（SKILL-INTERVIEW.md）：技能盘点、STAR故事库、模拟面试、公司调研
- 薪资谈判 Skill（SKILL-SALARY.md）：市场调研、反报价策略、Offer比较矩阵
- 求职工具包 Skill（SKILL-TOOLKIT.md）：申请表填写、投递跟进、作品集撰写、跨行求职
- HR简历审核（hr-resume-reviewer.md）：AI痕迹识别、数据可信度审查
- 投递跟踪（career-tracker.md）：BOSS直聘投递记录、跟进节奏、统计面板
- 简历生成脚本（generate_resume.py）：PDF/Word读取、HTML生成、证件照嵌入
- 参考文档（references/）：JD分析规则、STAR方法论、写作规则、ATS格式、证据链、真相边界、面试拷问、BOSS打招呼话术
- 合并版 Prompt（PROMPT-MERGED.md）：精简版，可在其他AI工具中使用
- 一键初始化（generate_resume.py setup）：依赖检查、模板生成
- 核心优势模块：HTML简历支持独立核心优势展示区
- 面试复盘模板：结构化反思面试表现
- 国内薪资谈判专章：13/14薪、五险一金、应届生策略、中文话术
- 完整触发词映射表：覆盖所有Skill的所有触发词
- GitHub 配置：Issue 模板、PR 模板、LICENSE