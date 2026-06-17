# 市场调研自动化工具套件
## 完整使用说明（README）
## 版本：v1.0 | McKinsey 高管汇报级别 | 六型框架全覆盖

---

## 一、套件文件清单

```
market-research-final/
│
├── 00_内部数据输入模板.md      ← 每次调研必填（你的核心输入）
├── market_research_tool.py    ← Python自动化脚本（有API版）
├── market_research_tool.mjs   ← Node.js自动化脚本（有API版）
├── claude_system_prompt.md    ← Claude.ai Projects配置（无API版）
├── coze_config.md             ← Coze平台完整配置文档
└── README.md                  ← 本文件
```

---

## 二、四种使用方式对比

| 使用方式 | 需要API | 技术门槛 | 自动化程度 | 输出格式 | 推荐场景 |
|---------|---------|---------|---------|---------|---------|
| **方式A**：Python脚本 | ✅ 需要 | 低（pip install）| 全自动 | Word + PPT + JSON | 日常工作，最推荐 |
| **方式B**：Node.js脚本 | ✅ 需要 | 低（npm install）| 全自动 | Word + PPT + JSON | 前端开发环境 |
| **方式C**：Claude.ai Projects | ❌ 不需要 | 零门槛 | 半自动（需复制粘贴）| Markdown（手动导出）| 临时调研，快速使用 |
| **方式D**：Coze平台 | ❌ 不需要 | 中（需配置Workflow）| 全自动 | Markdown | 团队共享，多人使用 |

---

## 三、方式A：Python脚本使用说明（推荐）

### 环境准备（一次性安装）

```bash
# 1. 确认Python版本 >= 3.9
python --version

# 2. 安装依赖包
pip install anthropic python-docx python-pptx

# 3. 设置API Key（二选一）
# 方法1：设置环境变量（推荐，更安全）
export ANTHROPIC_API_KEY="your-api-key-here"   # Mac/Linux
set ANTHROPIC_API_KEY=your-api-key-here         # Windows

# 方法2：直接填写在脚本CONFIG中（不推荐用于生产环境）
# 打开 market_research_tool.py，找到 CONFIG 区域
# 将 "your-api-key-here" 替换为你的真实API Key
```

### 每次调研的操作流程

**Step 1：填写内部数据输入模板**

打开 `00_内部数据输入模板.md`，填写以下必填项：

```
【决策问题】→ 领导要做的决策（一句话）
【产品/服务名称】→ 产品名称
【核心技术指标】→ 关键技术参数（越具体越好）
【产品发挥优势的必要条件】→ 在什么环境/场景下有优势
【候选市场/目标市场】→ 需要对比的市场
【目标客户类型】→ 客户类型假设
【公司资源约束】→ 当前能做什么、不能做什么
```

根据调研类型，勾选对应类型并填写「类型专属区域」（Section 4）。

**Step 2：确认脚本配置**

打开 `market_research_tool.py`，确认 `CONFIG` 区域：

```python
CONFIG = {
    "api_key": "your-api-key",           # 或使用环境变量
    "input_file": "00_内部数据输入模板.md",  # 输入文件路径
    "output_dir": "output_reports",        # 输出目录
    "enable_web_search": True,             # 是否开启自动网络搜索
}
```

**Step 3：运行脚本**

```bash
python market_research_tool.py
```

**Step 4：查看输出文件**

脚本运行完成后，在 `output_reports/` 目录找到：
- `市场调研报告_[产品名]_[日期时间].docx`  ← Word报告（直接提交给领导）
- `市场调研报告_[产品名]_[日期时间].pptx`  ← PPT汇报（演示用）
- `raw_output_[产品名]_[日期时间].json`     ← 原始内容（用于二次加工）

### 预计运行时间

| 调研类型 | 模块数量 | 预计时间 |
|---------|---------|---------|
| A型（新市场进入） | 9个模块 | 8-12分钟 |
| B型（竞争情报） | 6个模块 | 5-8分钟 |
| C型（客户需求） | 6个模块 | 5-8分钟 |
| D型（市场深耕） | 6个模块 | 5-8分钟 |
| E型（产品定价） | 6个模块 | 5-8分钟 |
| F型（品牌传播） | 6个模块 | 5-8分钟 |

---

## 四、方式B：Node.js脚本使用说明

### 环境准备

```bash
# 1. 确认Node.js版本 >= 18
node --version

# 2. 安装依赖
npm install @anthropic-ai/sdk docx pptxgenjs

# 3. 设置API Key
export ANTHROPIC_API_KEY="your-api-key-here"
```

### 运行

```bash
node market_research_tool.mjs
```

输出文件与Python版完全相同。

---

## 五、方式C：Claude.ai Projects 使用说明（无API）

### 一次性配置（5分钟）

1. 打开 [claude.ai](https://claude.ai)，登录账号
2. 点击左侧「Projects」→「New Project」
3. 点击「Project Instructions」
4. 打开 `claude_system_prompt.md`，复制「=== 开始 ===」到「=== 结束 ===」之间的全部内容
5. 粘贴到 Project Instructions 输入框，点击保存

配置完成后，此 Project 就是你的专属市场调研助手，永久生效。

### 每次调研的操作流程

**Step 1：填写内部数据模板**
打开 `00_内部数据输入模板.md`，填写好所有必填项。

**Step 2：发送给Claude**
在 Project 对话框中，将填写好的模板内容完整粘贴进去，发送。

**Step 3：Claude自动执行**
Claude会自动：
- 判断调研类型
- 逐模块生成分析内容
- 引用权威数据来源
- 执行质检清单

**Step 4：导出为Word/PPT**
- **Word**：点击对话框右上角「Copy」→ 粘贴到Word → 调整格式
- **PPT**：使用以下提示词让Claude帮你生成PPT大纲，再手动制作：
  > 「请将上述报告内容整理为PPT大纲，每页幻灯片给出标题+3-5个要点」

### 快捷指令

在Project对话框中输入以下指令触发对应功能：

| 指令 | 功能 |
|------|------|
| `/template` | 输出内部数据输入模板引导 |
| `/detect` | 自动判断调研类型 |
| `/full` | 生成完整报告（所有模块） |
| `/exec` | 只生成执行摘要+决策矩阵（快速版，约5分钟）|
| `/module 市场筛选` | 只生成指定模块 |
| `/check` | 对已有报告进行质检 |
| `/sources` | 输出当前调研类型的权威数据来源清单 |
| `/battlecard 竞品名` | 快速生成针对某竞品的销售Battlecard |

---

## 六、方式D：Coze 平台使用说明

详细配置见 `coze_config.md`，以下是快速启动步骤：

1. 登录 [coze.com](https://coze.com)
2. 「创建 Bot」→ 填写名称和描述
3. 将 `coze_config.md` 第三节的 System Prompt 粘贴到「人设与回复逻辑」
4. 在「插件」中启用：**Web Search**（必须）+ **Link Reader**（推荐）
5. 按 `coze_config.md` 第五节的 Workflow 图搭建自动化流程
6. 在「知识库」中上传 `00_内部数据输入模板.md` 和产品文档
7. 发布为团队Bot，所有人共享使用

Coze版本的优势：
- 无需API Key，使用Coze平台内置额度
- 支持团队多人共享
- 可以集成到企业微信/飞书/Slack等

---

## 七、六种调研类型快速判断

每次开始调研前，先用这个表格判断类型：

| 领导的核心问题 | 调研类型 | 核心输出 |
|-------------|---------|---------|
| 应该进入哪个新市场？怎么进？ | **A型：新市场进入** | 市场优先级排序+GTM计划 |
| 竞品在做什么？我们如何反制？ | **B型：竞争情报** | 威胁评级+Battlecard |
| 客户真正要什么？愿意为什么付钱？ | **C型：客户需求** | 需求优先级+ICP更新 |
| 在现有市场如何扩大份额？ | **D型：市场深耕** | 增长杠杆ROI排序 |
| 产品该迭代什么？价格怎么定？ | **E型：产品定价** | 路线图优先级+定价建议 |
| 品牌认知是什么？信息有效吗？ | **F型：品牌传播** | 认知差距+渠道预算 |

---

## 八、数据来源权威性标准

工具会优先引用以下来源（按权威性排序）：

**L1级（最高权威，优先引用）**
- World Bank（worldbank.org）：宏观经济、营商环境、发展指标
- IMF（imf.org）：经济预测、汇率
- IEA（iea.org）：能源政策、装机数据
- IRENA（irena.org）：可再生能源成本趋势
- McKinsey Global Institute（mckinsey.com）：行业深度报告

**L2级（高权威，常规引用）**
- BloombergNEF（bnef.com）：储能、光伏市场数据
- Gartner（gartner.com）：技术采用趋势
- Wood Mackenzie：能源市场分析
- Reuters / Bloomberg / FT：市场新闻

**工具自动执行的数据来源验证**：
- 每条数据附来源机构+年份
- L4级数据自动标注「[待验证]」
- 超过24个月的市场规模数据不被引用
- 单一来源的份额数据会提示需要交叉验证

---

## 九、常见问题解答

**Q：脚本运行报错「ModuleNotFoundError」**
```bash
# 重新安装依赖
pip install --upgrade anthropic python-docx python-pptx
```

**Q：API Key 报错「invalid_api_key」**
- 确认API Key是否已激活（去console.anthropic.com确认）
- 确认是否有账户余额
- 确认Key是否以「sk-ant-」开头

**Q：生成的数据没有引用来源**
- 检查 `enable_web_search` 是否设置为 `True`
- 网络搜索功能需要API账户支持工具调用，确认账户权限

**Q：Word文档格式有问题**
- 用Microsoft Word打开（不要用WPS），格式会更准确
- 如果需要调整样式，在Word中「设计」→选择你喜欢的主题

**Q：PPT内容太多，一页放不下**
- PPT默认每页提取前9个要点
- 可以在Python脚本中修改：`bulletLines = lines.slice(0, 9)` 改为更少的数字
- 建议配合PPT文案手动精简，PPT是辅助工具

**Q：如何更新到最新版本**
- 关注工具的更新说明
- 新版本会保持相同的模板格式，直接替换脚本文件即可

---

## 十、输出质量自检清单

**提交报告前，逐项检查：**

**结构层面**
- [ ] 执行摘要5分钟内可读完，领导读完知道下一步做什么
- [ ] 各市场分析深度与推荐优先级匹配（重要市场分析更深）

**内容层面**
- [ ] 每个关键数据有来源标注（机构+年份）
- [ ] 每个「我们的优势」有数据支撑（非泛泛描述）
- [ ] 客户描述达到ICP级别（行业+规模+地理+场景）
- [ ] 竞品分析基于具体场景的胜负（非静态参数表）
- [ ] TAM/SAM/SOM有逻辑推导（非直接引用他人数字）

**决策层面**
- [ ] 市场优先级结论明确（无「各有优势」等模糊表述）
- [ ] 第一步行动具体（时间+动作+对象+量化结果）
- [ ] 风险有应对方式（预警信号+预防+止损触发点）

**数据层面**
- [ ] 同一指标全文口径一致
- [ ] 无超过24个月的市场规模数据
- [ ] L4级数据标注[待验证]


---

*README版本：v1.0 | 工具版本：v1.0 | 最后更新：2026年*
