/**
 * 市场调研自动化工具 — Node.js 完整版
 * =====================================================
 * 功能：
 *   1. 读取内部数据输入模板（用户填写的信息）
 *   2. 自动判断调研类型（A/B/C/D/E/F）
 *   3. 针对每个模块自动调用 Claude API（含网络搜索）
 *   4. 生成 McKinsey 级别分析内容
 *   5. 自动生成 Word 报告（.docx）
 *   6. 自动生成 PowerPoint 汇报（.pptx）
 *   7. 自动保存 JSON 中间结果
 *
 * 安装依赖：
 *   npm install @anthropic-ai/sdk docx pptxgenjs fs-extra chalk
 *
 * 运行：
 *   node market_research_tool.mjs
 *   或：ANTHROPIC_API_KEY=your-key node market_research_tool.mjs
 */

import Anthropic from "@anthropic-ai/sdk";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  HeadingLevel,
  AlignmentType,
  WidthType,
  ShadingType,
  BorderStyle,
  PageBreak,
} from "docx";
import PptxGenJS from "pptxgenjs";
import fs from "fs";
import path from "path";

// ─────────────────────────────────────────────────────────────
// CONFIG：每次调研只需要修改这里
// ─────────────────────────────────────────────────────────────

const CONFIG = {
  apiKey: process.env.ANTHROPIC_API_KEY || "your-api-key-here",
  inputFile: "00_内部数据输入模板.md",
  outputDir: "output_reports",
  model: "claude-sonnet-4-20250514",
  maxTokens: 1500,
  enableWebSearch: true,
  language: "中文",
};

// ─────────────────────────────────────────────────────────────
// 权威数据来源层级（嵌入所有提示词）
// ─────────────────────────────────────────────────────────────

const AUTHORITY_SOURCES = `
## 数据来源权威性要求（严格遵守）

### L1级（最高权威，优先引用）
- 国际机构：World Bank, IMF, IEA, IRENA, WTO, UN, OECD
- 政府官方：各国政府统计局、能源部、财政部
- 顶级咨询：McKinsey Global Institute, BCG, Bain, Deloitte, PwC, EY
- 金融监管：美联储, ECB, BIS

### L2级（高权威，常规引用）
- 顶级行业研究：BloombergNEF, Gartner, Forrester, IDC, Wood Mackenzie
- 权威媒体：路透社, 彭博社, 《金融时报》, 《经济学人》, WSJ
- 行业协会：SEIA, BNEF, 各国光伏协会, 矿业协会

### L3级（中等，须交叉验证）
- 上市公司年报、可持续发展报告
- PV Tech, Renew Economy, Mining Technology

### L4级（参考，必须标注[待验证]）
- Statista, IBISWorld, Grand View Research
- LinkedIn, Crunchbase

### 禁止引用
- 无来源数据、AI生成内容、超过24个月市场规模数据
`;

// ─────────────────────────────────────────────────────────────
// 六种调研类型的模块定义
// ─────────────────────────────────────────────────────────────

const TYPE_MODULES = {
  A: {
    name: "新市场进入研究（Market Entry Research）",
    modules: [
      ["exec_summary", "执行摘要"],
      ["product_anchor", "产品定位锚定"],
      ["market_screening", "市场筛选评分矩阵"],
      ["market_deep_dive", "重点市场深度分析"],
      ["icp", "目标客户画像（ICP）"],
      ["competitive_winrate", "竞争格局与差异化胜率"],
      ["business_model", "商业模式设计"],
      ["gtm_plan", "市场进入路径（GTM）"],
      ["decision_matrix", "优先级决策矩阵与最终结论"],
    ],
  },
  B: {
    name: "竞争情报研究（Competitive Intelligence）",
    modules: [
      ["exec_summary", "执行摘要"],
      ["competitive_landscape", "竞争格局总览"],
      ["competitor_profile", "竞品深度画像"],
      ["winrate_analysis", "差异化胜率分析"],
      ["battlecard", "销售Battlecard"],
      ["counter_plan", "反制行动计划"],
    ],
  },
  C: {
    name: "客户需求研究（Customer & Demand Insight）",
    modules: [
      ["exec_summary", "执行摘要"],
      ["research_design", "研究设计与访谈框架"],
      ["demand_analysis", "需求优先级矩阵"],
      ["icp_update", "客户画像精确化"],
      ["willingness_to_pay", "支付意愿与定价洞察"],
      ["gtm_message", "GTM信息优化建议"],
    ],
  },
  D: {
    name: "市场深耕研究（Market Penetration）",
    modules: [
      ["exec_summary", "执行摘要"],
      ["growth_diagnosis", "增长结构诊断"],
      ["customer_health", "客户健康度分析"],
      ["untapped_market", "渗透率空白识别"],
      ["growth_levers", "增长杠杆ROI排序"],
      ["action_plan", "季度行动计划"],
    ],
  },
  E: {
    name: "产品与定价研究（Product & Pricing）",
    modules: [
      ["exec_summary", "执行摘要"],
      ["product_assessment", "产品功能价值评估"],
      ["roadmap_priority", "产品路线图优先级矩阵"],
      ["pricing_diagnosis", "定价现状诊断"],
      ["willingness_to_pay_e", "支付意愿研究"],
      ["pricing_structure", "定价结构与实施路径"],
    ],
  },
  F: {
    name: "品牌与传播研究（Brand & Communication）",
    modules: [
      ["exec_summary", "执行摘要"],
      ["brand_awareness", "品牌认知漏斗分析"],
      ["brand_association", "品牌联想与竞品对比"],
      ["message_effectiveness", "核心信息有效性"],
      ["channel_efficiency", "渠道触达效率分析"],
      ["brand_strategy", "品牌战略与90天行动"],
    ],
  },
};

// ─────────────────────────────────────────────────────────────
// 调研类型自动判断
// ─────────────────────────────────────────────────────────────

function detectResearchType(content) {
  if (content.includes("[x] A型") || content.includes("[X] A型")) return "A";
  if (content.includes("[x] B型") || content.includes("[X] B型")) return "B";
  if (content.includes("[x] C型") || content.includes("[X] C型")) return "C";
  if (content.includes("[x] D型") || content.includes("[X] D型")) return "D";
  if (content.includes("[x] E型") || content.includes("[X] E型")) return "E";
  if (content.includes("[x] F型") || content.includes("[X] F型")) return "F";

  if (content.includes("新市场") || content.includes("进入") || content.includes("候选市场")) return "A";
  if (content.includes("竞品") || content.includes("竞争情报") || content.includes("battlecard")) return "B";
  if (content.includes("客户需求") || content.includes("用户访谈") || content.includes("支付意愿")) return "C";
  if (content.includes("市场深耕") || content.includes("增长") || content.includes("流失率")) return "D";
  if (content.includes("产品定价") || content.includes("路线图") || content.includes("定价策略")) return "E";
  if (content.includes("品牌") || content.includes("传播") || content.includes("认知度")) return "F";

  return "A";
}

// ─────────────────────────────────────────────────────────────
// 提取内部数据摘要
// ─────────────────────────────────────────────────────────────

function extractContextSummary(content) {
  const summary = {};
  const patterns = {
    product: /【产品\/服务名称】\s*\n→\s*(.+)/,
    markets: /【候选市场\/目标市场】\s*\n→\s*(.+)/,
    decision: /【决策问题】\s*\n→\s*(.+)/,
  };
  for (const [key, pattern] of Object.entries(patterns)) {
    const match = content.match(pattern);
    if (match) summary[key] = match[1].trim();
  }
  return summary;
}

// ─────────────────────────────────────────────────────────────
// 系统提示词构建
// ─────────────────────────────────────────────────────────────

function buildSystemPrompt(researchType) {
  const typeInfo = TYPE_MODULES[researchType];
  return `你是麦肯锡资深市场策略顾问，专精B2B工业/能源/技术产品的高管汇报级市场调研。

当前调研类型：${typeInfo.name}

${AUTHORITY_SOURCES}

## 输出质量红线（违反则返工）
1. 结论必须明确：禁止「各有优劣」「视情况而定」「需进一步研究」
2. 每个洞察完成四层：What（观察）→ Why（原因）→ So What（含义）→ Now What（行动）
3. 数据必须有来源：每个关键数据附来源机构+年份，L4级标注[待验证]
4. 客户描述ICP级别：行业+规模+地理+场景，禁止泛化
5. 竞品分析场景化：在[具体条件]下，[竞品]问题[数据]，我们能做到[数据]，客户收益[可量化]
6. 行动必须可执行：时间节点+具体动作+负责方+量化成功标准
7. 风险有应对：预警信号+预防措施+止损触发点
8. 口径全文一致：同一指标只用一个单位

## 叙事结构
金字塔原则：结论先行 → 核心论点 → 支撑数据 → 具体行动
输出格式：Markdown，使用##和###标题层级`;
}

// ─────────────────────────────────────────────────────────────
// 各模块提示词
// ─────────────────────────────────────────────────────────────

function buildModulePrompt(moduleKey, context) {
  const prompts = {
    exec_summary: `基于以下调研背景，生成「执行摘要」。
要求：不超过500字，高管5分钟内读完可直接拍板，Markdown格式。

${context}

必须包含：
## 核心结论
推荐方向（1-2个）+ 3条数据支撑理由（每条引用权威数据来源）

## 次选与暂缓
各一句话，原因具体，附数据

## 第一步行动
[具体时间节点] → [做什么] → [找谁] → [预期可量化结果]

## 90天里程碑
3个可量化检查点，含负责方`,

    product_anchor: `基于以下信息，生成「产品定位锚定」（整份报告的技术基准）。

${context}

必须包含：
## 核心技术指标对比表（我们 vs 主要竞品，每个指标附权威数据来源）

## 赢的场景（2-3个）
格式：场景条件→竞品量化问题[来源]→我们量化优势[来源]→客户财务收益[具体数字]

## 输的场景（1-2个，诚实写出，含对市场筛选的含义）

## 权重建议（对后续市场评分矩阵哪个维度应设最高权重，理由）`,

    market_screening: `基于以下信息，生成「市场筛选评分矩阵」。

${context}

六维度评分（1-5分）：需求强度25%、技术匹配25%、支付能力20%、进入难度15%、竞争强度10%、战略价值5%

每个分数必须有具体数据支撑（来源标注），禁止凭印象打分。

## 评分矩阵表格（各市场×各维度，含加权总分，每格附一句话依据+数据来源）

## 市场分层结论
立即进入（>4.0）/观察跟进（3.0-4.0）/暂缓（<3.0），每层3条理由附数据来源

## 唯一最优市场（如果只能做一个，推荐哪个，附3条核心理由，必须明确）

## 暂缓市场重启条件`,

    market_deep_dive: `基于以下信息，对评分最高市场生成「深度分析」。

${context}

## 市场本质（一句话锚点）
格式：「[市场] = [客户核心痛点] × [现有方案结构性缺陷] × [我们独特价值] 的交叉市场」

## 供需结构与三大核心矛盾
每个矛盾：现状（数据+L1/L2来源）→ 本质问题 → 对我们的具体机会

## 市场规模（TAM→SAM→SOM，每层有推导逻辑，SOM基于接触量×转化率×客户价值）

## 标杆案例（1-2个）：成功因素+局限性[有数据]+我们的超越机会

## 风险矩阵（概率×影响→预警信号→预防→止损触发点）`,

    icp: `基于以下信息，生成「目标客户画像（ICP）」。
定义：最可能第一批成交且能产生示范效应的客户类型。

${context}

## 客户特征（ICP级别：行业+子行业+规模量化+具体地理位置城市级+组织特征）

## 核心痛点（3个）
格式：现状（客户原话/公开资料引用）→ 量化损失 → 我们能解决的程度

## 购买触发器（2-3个具体事件场景）

## 决策链表格（角色/职位/决策作用/核心关注点/最担心的风险）

## 供应商评估标准（按重要性排序，5-7条）

## 典型代表企业（3-5家，附符合ICP的具体依据）

## 触达路径（渠道+时机+切入话题，禁止写「介绍产品」）`,

    competitive_winrate: `基于以下信息，生成「竞争格局与差异化胜率」。

${context}

## 竞争阵营分析表（竞争类型/代表企业/市场份额[来源]/核心弱点[场景化]/非对称优势）

## 三场景差异化胜率（严格格式）
场景①：在[具体条件]下，[竞品]问题[量化数据+来源]，我们能做到[量化数据+来源]，客户收益[可量化财务]

## 应主动避开的场景（结构性劣势）

## 差异化一句话定位：「[我们]是[ICP]在[具体场景]下比[竞品]更好的[核心价值]方案」

## 非对称优势可持续性（竞品复制需要多长时间，为什么）`,

    business_model: `基于以下信息，生成「商业模式设计」。

${context}

## 模式适用性对比表（设备直销/EPC/PPA/EaaS/合资，含适配客户/资金要求/客户门槛/首单难度/规模化潜力）

## 第一阶段推荐模式（必须明确，附推荐理由3条[有数据]+关键成功要素+合同结构要点+首单规模区间）

## 第二阶段演进方向（12-36个月）

## 必须提前准备的前提条件清单（认证/合作伙伴/融资，每项附时间估算）`,

    gtm_plan: `基于以下信息，生成「市场进入路径（GTM）」。

${context}

## 进入路径选择与推荐（直接/合作伙伴/展会/标杆项目，推荐明确+理由）

## 客户触达策略（ICP在哪里获取信息+如何建立第一次有效接触+最佳切入话题）

## 分阶段行动计划
第一阶段（0-6月）：核心目标+具体动作[做什么/找谁/方式/预期结果]+量化成功标准
第二阶段（6-18月）：同格式
第三阶段（18月+）：规模复制逻辑

## 启动前必须准备清单`,

    decision_matrix: `基于以下信息，生成「优先级决策矩阵与最终结论」（报告最重要的一页）。

${context}

## 市场优先级决策矩阵（机会评级/进入时机/推荐模式/关键前提条件/风险等级）

## 最终战略结论（必须明确，绝不模糊）
第一优先：[市场]，理由①②③（每条含权威数据来源）
第二梯队：[市场]，进入条件
暂缓：[市场]，具体原因（一句话）

## 第一步行动（最终落地）
[时间] [具体动作] [找谁/通过什么渠道] [预期可量化结果]

## 90天里程碑（时间点/具体里程碑/量化成功标准/负责方）`,

    competitive_landscape: `基于以下信息，生成「竞争格局总览」（B型）。

${context}

## 竞争阵营三层分类（直接竞品/间接竞品/潜在进入者，含市场份额[来源]+威胁等级+监控优先级）

## 市场份额动态（过去12个月变化+驱动因素+我们相对地位趋势[均需来源]）

## 最高威胁竞品定性（具体威胁场景，非笼统「综合竞争力强」）

## 竞争格局未来12个月预判（更激烈/分散，驱动因素）

## 资源分配决策含义`,

    competitor_profile: `基于以下信息，生成「竞品深度画像」（B型）。

${context}

对每个主要竞品：
## 战略意图分析（融资/招聘/新市场/产品信号→战略推断→置信度H/M/L）

## 产品能力场景化对比（禁止纯参数表，每个对比绑定具体使用场景，附第三方来源）

## 四类弱点深挖（客户公开抱怨[引用原文]/技术局限[有数据]/服务弱点/覆盖空白）

## 未来6个月最可能重大动作预判`,

    winrate_analysis: `基于以下信息，生成「差异化胜率分析」（B型）。

${context}

## 竞争定位地图（按ICP最在意的2个维度[非技术维度]定位各竞品）

## 三场景差异化胜率
格式：在[条件]下，[竞品]问题[数据+来源]，我们能做到[数据+来源]，客户收益[可量化]

## 应避开的竞争场景（结构性弱点）

## 差异化优势可持续性评估`,

    battlecard: `基于以下信息，生成「销售Battlecard」（B型，单页浓缩）。

${context}

## 一句话定位差异

## 我们的3个核心优势（每个附数据+客户可量化价值）

## 竞品声称 vs 我们的应对话术 vs 支撑证据（表格）

## 客户常见异议 vs 专业回应 vs 关键证据（表格）

## 主动推的场景 vs 需要转移的场景`,

    counter_plan: `基于以下信息，生成「反制行动计划」（B型）。

${context}

## 竞品近期重大动作的即时应对（按H/M/L威胁等级，含时间节点和负责方）

## 主动进攻计划（从竞品手中赢得客户：识别标准+切入时机+切入话术）

## CI持续监控机制（月度监控清单+季度深度更新触发条件）`,

    research_design: `基于以下信息，生成「研究设计与访谈框架」（C型）。

${context}

## 访谈对象分层设计（现有客户/流失客户/竞品客户/非客户，各类型数量+找到方式）

## 完整访谈问题框架（现状探索/痛点深挖/解决方案评估/Van Westendorp四问法）

## 三层需求识别框架（表达需求/真实需求/潜在需求的区别与识别方法）`,

    demand_analysis: `基于以下信息，生成「需求优先级矩阵」（C型）。

${context}

## 需求优先级评分矩阵（提及频率×痛感强度×支付意愿×当前满足度，综合优先级计算）

## 四象限分类（立即解决/保持优势/观察跟进/过度投资）

## 最被低估的需求（高优先级但当前未解决）

## 不值得追逐的需求（表达强烈但支付意愿低）

## 产品迭代决策含义`,

    icp_update: `基于以下信息，生成「客户画像精确化」（C型）。

${context}

## ICP更新对比表（原有假设 vs 研究发现 vs 更新后定义）

## 高价值 vs 低价值客户分化特征

## 需要新增的ICP细分

## 需要排除的客户画像（与能力不匹配）

## 销售资质审核标准更新建议`,

    willingness_to_pay: `基于以下信息，生成「支付意愿与定价洞察」（C型）。

${context}

## Van Westendorp四价格点（拒绝价/可接受上限/便宜购买/怀疑价+建议区间）

## 分客户群支付意愿差异

## 当前定价位置诊断（是否存在价值低估）

## 溢价空间分析（为哪些特性，幅度多大）

## 定价调整建议与风险`,

    gtm_message: `基于以下信息，生成「GTM信息优化建议」（C型）。

${context}

## 客户语言 vs 我们的语言对比表

## 核心价值主张重构（基于客户语言）

## 分决策链角色的差异化信息

## 最有力的3个客户引言（可直接用于营销）`,

    growth_diagnosis: `基于以下信息，生成「增长结构诊断」（D型）。

${context}

## 增长来源分解（新客获取/现有客扩展/价格调整/市场自然增长，各占比+趋势+数据来源）

## 增长漏出分析（客户流失/降级/未转化线索，各占比+根本原因）

## 净增长真实来源（「我们做好了」vs「市场在涨」的分离）

## 市场份额动态（相对竞品地位变化，非绝对数字）

## 增长质量评估与战略含义`,

    customer_health: `基于以下信息，生成「客户健康度分析」（D型）。

${context}

## 客户价值分层（A/B/C/D级定义+各层客户数+ARR占比+策略）

## 健康度评分模型（使用频率/续约意愿/NPS/扩展行为/服务质量，权重）

## 流失风险前10名识别（干预计划）

## 扩展机会前10名识别（最容易的增收来源）

## 流失根因分析（可阻止 vs 不可阻止+比例+应对）`,

    untapped_market: `基于以下信息，生成「渗透率空白识别」（D型）。

${context}

## 细分市场渗透率矩阵（按行业/规模/地区/场景，潜在客户数+当前覆盖+渗透率+未覆盖原因）

## 最大渗透率空白及进入障碍分析

## 「一个改变解锁多个客户群」的机会

## ROI最高的空白攻克路径`,

    growth_levers: `基于以下信息，生成「增长杠杆ROI排序」（D型）。

${context}

## 增长杠杆全景（含预期贡献+所需投入+ROI）

## 每个杠杆四层分析（What量化→Why为什么有效→So What不做的损失→Now What第一步）

## 最终优先级决策表（按ROI，含启动时间和负责方）

## 资源分配建议（防御%/进攻%/探索%+理由）`,

    action_plan: `基于以下信息，生成「季度行动计划」（D型）。

${context}

## 90天执行路线图（第1/2/3个月：核心目标+具体动作+量化成功标准）

## 关键里程碑表（成功标准/时间/负责方）

## 风险预案（每个主要行动的风险+应急措施）`,

    product_assessment: `基于以下信息，生成「产品功能价值评估」（E型）。

${context}

## Kano模型功能分类（必备/期望/魅力/无差异，当前满足度+竞品对比+数据来源）

## 被错误投入资源的功能（无差异功能清单，释放研发产能建议）

## 被低估的魅力功能（未宣传的差异化资产）

## 产品使用数据分析（核心功能使用率/Aha Moment/流失预测信号/Time-to-Value）

## 迭代资源分配建议`,

    roadmap_priority: `基于以下信息，生成「产品路线图优先级矩阵」（E型）。

${context}

## 需求来源汇总（客户/败单/流失/竞品差距，真实性验证）

## 优先级评分矩阵（市场需求强度×收入影响×战略价值×实现复杂度×竞争紧迫性）

## 四象限决策（立即做/计划做/再议/不做）

## 「不做」决策及明确理由

## Quick Win识别（2-4周高价值改善）

## 路线图与定价联动逻辑`,

    pricing_diagnosis: `基于以下信息，生成「定价现状诊断」（E型）。

${context}

## 价值-价格对齐分析（每个功能的客户感知价值 vs 是否计费 → 价值捕获状态）

## 最大价值捕获缺口（免费提供的高价值功能）

## 过度定价风险（单独计费但感知低的功能）

## 竞品定价基准对比

## 定价溢价的合理空间`,

    willingness_to_pay_e: `基于以下信息，生成「支付意愿研究」（E型）。

${context}

## Van Westendorp四价格点（基于真实访谈数据，n≥15）

## 分客户层支付意愿差异

## 当前定价位置诊断

## 价格弹性评估（涨价10/20/30%的预估流失率）

## 定价分层机会`,

    pricing_structure: `基于以下信息，生成「定价结构与实施路径」（E型）。

${context}

## 定价结构选项评估（按席位/用量/功能分层/成果分成/混合）

## 最优定价结构推荐（必须明确，附与客户价值感知的对齐逻辑）

## 定价分层设计（分割点依据）

## 涨价实施路径（价值强化→新客试行→现有客通知→全面执行→效果评估）

## 风险缓冲方案`,

    brand_awareness: `基于以下信息，生成「品牌认知漏斗分析」（F型）。

${context}

## 品牌认知五层漏斗（知名度/理解度/差异化认知/可信度/偏好度，当前数据+目标+差距）

## 最大漏斗断层识别与商业影响量化

## 与竞品的对比（竞品在哪层领先）

## 品牌投入方向决策含义`,

    brand_association: `基于以下信息，生成「品牌联想与竞品对比感知」（F型）。

${context}

## 品牌联想词云分析（客户实际高频词 vs 我们希望的 vs 差距）

## 「品牌现实」vs「品牌意图」最大差距

## 意外的正面联想（被忽视的差异化资产）

## 需要纠正的负面联想

## 竞品品牌对比感知矩阵`,

    message_effectiveness: `基于以下信息，生成「核心信息有效性」（F型）。

${context}

## 当前核心信息测试结果（理解率/共鸣率/可信度/行动意愿）

## 高共鸣 vs 无效信息分类（附客户原话）

## 「最有力的一句话」

## 「技术语言陷阱」识别

## 核心价值主张更新建议（分角色）`,

    channel_efficiency: `基于以下信息，生成「渠道触达效率分析」（F型）。

${context}

## 目标受众信息接触点调研（使用频率/信任度/我们存在感/投入产出）

## 最高价值渠道（高信任度+我们低存在感）

## 资源错配渠道（高投入但低受众覆盖）

## 「品牌沙漠」识别

## 下年度营销预算分配建议`,

    brand_strategy: `基于以下信息，生成「品牌战略与90天行动」（F型）。

${context}

## 品牌定位声明（内部格式）
对于[ICP]，面对[痛点]，[品牌]是[类别]，提供[差异化价值]，因为[独特能力]，不同于[竞品]

## 三层信息体系（核心主张/分层信息/证明层）

## 90天行动计划（每步有量化成功标准）

## 品牌投资ROI预期`,
  };

  return (
    prompts[moduleKey] ||
    `基于以下信息，生成「${moduleKey}」模块，严格遵守麦肯锡高管汇报质量标准。\n\n${context}`
  );
}

// ─────────────────────────────────────────────────────────────
// Claude API 调用（含网络搜索）
// ─────────────────────────────────────────────────────────────

async function callClaude(client, prompt, system, cfg) {
  const params = {
    model: cfg.model,
    max_tokens: cfg.maxTokens,
    system,
    messages: [{ role: "user", content: prompt }],
  };

  if (cfg.enableWebSearch) {
    params.tools = [{ type: "web_search_20250305", name: "web_search" }];
  }

  try {
    const msg = await client.messages.create(params);
    const textParts = msg.content
      .filter((b) => b.type === "text" && b.text)
      .map((b) => b.text);
    return textParts.join("\n");
  } catch (e) {
    return `[生成失败] ${e.message}`;
  }
}

// ─────────────────────────────────────────────────────────────
// Word 文档生成
// ─────────────────────────────────────────────────────────────

function parseMarkdownToDocxChildren(content) {
  const lines = content.split("\n");
  const children = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      children.push(new Paragraph({ text: "" }));
      continue;
    }

    if (trimmed.startsWith("## ")) {
      children.push(
        new Paragraph({
          text: trimmed.slice(3),
          heading: HeadingLevel.HEADING_2,
        })
      );
    } else if (trimmed.startsWith("### ")) {
      children.push(
        new Paragraph({
          text: trimmed.slice(4),
          heading: HeadingLevel.HEADING_3,
        })
      );
    } else if (trimmed.startsWith("- ") || trimmed.startsWith("* ")) {
      children.push(
        new Paragraph({
          text: trimmed.slice(2),
          bullet: { level: 0 },
        })
      );
    } else if (trimmed.startsWith("**") && trimmed.endsWith("**")) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({ text: trimmed.slice(2, -2), bold: true }),
          ],
        })
      );
    } else if (!trimmed.startsWith("|")) {
      children.push(new Paragraph({ text: trimmed }));
    }
  }

  return children;
}

async function generateWordReport(
  modulesOutput,
  moduleNames,
  researchType,
  contextSummary,
  outputPath
) {
  const typeInfo = TYPE_MODULES[researchType];
  const now = new Date();
  const dateStr = now.toLocaleDateString("zh-CN", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });

  const sections = [];

  // 封面
  sections.push({
    children: [
      new Paragraph({ text: "" }),
      new Paragraph({ text: "" }),
      new Paragraph({
        text: "市场调研报告",
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: typeInfo.name,
            size: 32,
            color: "4682B4",
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({ text: "" }),
      new Paragraph({
        children: [
          new TextRun({
            text: `产品：${contextSummary.product || "—"}  |  市场：${contextSummary.markets || "—"}  |  ${dateStr}`,
            size: 22,
            color: "666666",
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({ text: "" }),
      new Paragraph({
        children: [
          new TextRun({
            text: "McKinsey 高管汇报级别 | 数据来源：L1-L2 权威机构",
            size: 20,
            color: "888888",
            italics: true,
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
    ],
  });

  // 各模块
  for (const [key, name] of moduleNames) {
    if (!modulesOutput[key]) continue;

    const content = modulesOutput[key];
    const moduleChildren = [
      new Paragraph({
        text: name,
        heading: HeadingLevel.HEADING_1,
        pageBreakBefore: true,
      }),
      ...parseMarkdownToDocxChildren(content),
    ];

    sections.push({ children: moduleChildren });
  }

  // 数据声明
  sections.push({
    children: [
      new Paragraph({
        text: "数据质量声明",
        heading: HeadingLevel.HEADING_1,
        pageBreakBefore: true,
      }),
      new Paragraph({
        text: "本报告数据来源遵循权威性层级标准：",
      }),
      new Paragraph({
        text: "• L1级（最高）：World Bank · IMF · IEA · IRENA · McKinsey Global Institute",
        bullet: { level: 0 },
      }),
      new Paragraph({
        text: "• L2级（高）：BloombergNEF · Gartner · Reuters · Bloomberg · Wood Mackenzie",
        bullet: { level: 0 },
      }),
      new Paragraph({
        text: "• L3级（中）：企业年报 · 权威行业媒体（须交叉验证）",
        bullet: { level: 0 },
      }),
      new Paragraph({
        text: "• [待验证]标注：需要进一步交叉验证的数据",
        bullet: { level: 0 },
      }),
      new Paragraph({ text: "" }),
      new Paragraph({
        children: [
          new TextRun({
            text: `报告生成时间：${now.toLocaleString("zh-CN")}  |  标准：McKinsey 高管汇报级别`,
            italics: true,
            color: "888888",
          }),
        ],
      }),
    ],
  });

  const doc = new Document({ sections });
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.log(`✅ Word 报告已生成：${outputPath}`);
}

// ─────────────────────────────────────────────────────────────
// PowerPoint 生成
// ─────────────────────────────────────────────────────────────

async function generatePptxReport(
  modulesOutput,
  moduleNames,
  researchType,
  contextSummary,
  outputPath
) {
  const typeInfo = TYPE_MODULES[researchType];
  const prs = new PptxGenJS();

  prs.layout = "LAYOUT_WIDE";
  prs.author = "市场调研自动化工具 v1.0";
  prs.subject = typeInfo.name;

  const DARK_BLUE = "1B3A6B";
  const MED_BLUE = "4682B4";
  const LIGHT_BLUE = "D6E4F0";
  const WHITE = "FFFFFF";
  const DARK_GRAY = "333333";
  const ACCENT = "FF8C00";
  const LIGHT_GRAY = "F5F5F5";

  // 封面
  const cover = prs.addSlide();
  cover.background = { color: DARK_BLUE };
  cover.addShape(prs.ShapeType.rect, {
    x: 0, y: 5.8, w: "100%", h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT },
  });
  cover.addText("市场调研报告", {
    x: 0.5, y: 1.2, w: 12.3, h: 1.2,
    fontSize: 44, bold: true, color: WHITE, align: "center",
  });
  cover.addText(typeInfo.name, {
    x: 0.5, y: 2.8, w: 12.3, h: 0.8,
    fontSize: 20, color: LIGHT_BLUE, align: "center",
  });
  cover.addText(
    `产品：${contextSummary.product || "—"}   |   市场：${contextSummary.markets || "—"}   |   ${new Date().toLocaleDateString("zh-CN")}`,
    { x: 0.5, y: 4.0, w: 12.3, h: 0.6, fontSize: 13, color: WHITE, align: "center" }
  );
  cover.addText("McKinsey 高管汇报级别  |  数据来源：L1-L2 权威机构", {
    x: 0.5, y: 6.8, w: 12.3, h: 0.4, fontSize: 11, color: LIGHT_BLUE, align: "center",
  });

  // 目录
  const tocSlide = prs.addSlide();
  tocSlide.background = { color: WHITE };
  tocSlide.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: "100%", h: 1.3, fill: { color: DARK_BLUE }, line: { color: DARK_BLUE },
  });
  tocSlide.addText("目  录", {
    x: 0.5, y: 0.2, w: 12.3, h: 0.9, fontSize: 26, bold: true, color: WHITE,
  });

  const moduleList = [...moduleNames].filter(([k]) => modulesOutput[k]);
  const perCol = Math.ceil(moduleList.length / 2);
  moduleList.forEach(([key, name], idx) => {
    const col = Math.floor(idx / perCol);
    const row = idx % perCol;
    tocSlide.addText(`${(idx + 1).toString().padStart(2, "0")}.  ${name}`, {
      x: 0.5 + col * 6.5, y: 1.5 + row * 0.65, w: 6.0, h: 0.55,
      fontSize: 13, color: DARK_GRAY,
    });
  });

  // 各模块幻灯片
  moduleList.forEach(([key, name], idx) => {
    const content = modulesOutput[key] || "";
    const lines = content
      .split("\n")
      .map((l) => l.trim())
      .filter((l) => l && !l.startsWith("#") && !l.startsWith("|"));
    const bulletLines = lines.slice(0, 9);

    const slide = prs.addSlide();
    slide.background = { color: WHITE };

    // 顶部色条
    slide.addShape(prs.ShapeType.rect, {
      x: 0, y: 0, w: "100%", h: 1.3,
      fill: { color: DARK_BLUE }, line: { color: DARK_BLUE },
    });
    // 编号
    slide.addText(`${(idx + 1).toString().padStart(2, "0")}`, {
      x: 0.2, y: 0.15, w: 0.9, h: 0.9, fontSize: 28, bold: true, color: ACCENT, align: "center",
    });
    // 模块名
    slide.addText(name, {
      x: 1.2, y: 0.2, w: 11.0, h: 0.9, fontSize: 22, bold: true, color: WHITE,
    });

    // 内容
    if (bulletLines.length > 0) {
      const bulletText = bulletLines.map((l) => {
        const clean = l.replace(/^[-*•]\s*/, "").replace(/\*\*/g, "");
        return { text: "• " + clean, options: { breakLine: true } };
      });
      slide.addText(bulletText, {
        x: 0.5, y: 1.5, w: 12.3, h: 5.6,
        fontSize: 13, color: DARK_GRAY, valign: "top",
        paraSpaceAfter: 6,
      });
    }

    // 页码
    slide.addText(`${idx + 3} / ${moduleList.length + 2}`, {
      x: 11.5, y: 7.1, w: 1.5, h: 0.3,
      fontSize: 10, color: "AAAAAA", align: "right",
    });
  });

  // 结尾
  const endSlide = prs.addSlide();
  endSlide.background = { color: DARK_BLUE };
  endSlide.addShape(prs.ShapeType.rect, {
    x: 0, y: 3.5, w: "100%", h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT },
  });
  endSlide.addText("数据来源与质量声明", {
    x: 0.5, y: 1.2, w: 12.3, h: 1.0, fontSize: 28, bold: true, color: WHITE, align: "center",
  });
  endSlide.addText(
    "L1级：World Bank · IMF · IEA · IRENA · McKinsey Global Institute\n" +
    "L2级：BloombergNEF · Gartner · Reuters · Bloomberg\n" +
    "L3级：企业年报 · 权威行业媒体\n" +
    "[待验证] 标注：需进一步交叉验证的数据",
    { x: 1, y: 2.8, w: 11.3, h: 2.5, fontSize: 14, color: LIGHT_BLUE, align: "center", paraSpaceAfter: 10 }
  );

  await prs.writeFile({ fileName: outputPath });
  console.log(`✅ PPT 报告已生成：${outputPath}`);
}

// ─────────────────────────────────────────────────────────────
// 主流程
// ─────────────────────────────────────────────────────────────

async function runResearch(cfg) {
  console.log("\n" + "=".repeat(60));
  console.log("  市场调研自动化工具 v1.0");
  console.log("  McKinsey 高管汇报级别");
  console.log("=".repeat(60));

  // 读取输入文件
  if (!fs.existsSync(cfg.inputFile)) {
    console.error(`❌ 输入文件不存在：${cfg.inputFile}`);
    console.error("请先填写 00_内部数据输入模板.md 后再运行");
    process.exit(1);
  }

  const inputContent = fs.readFileSync(cfg.inputFile, "utf-8");
  console.log(`✅ 已读取输入文件：${cfg.inputFile}`);

  const contextSummary = extractContextSummary(inputContent);
  const researchType = detectResearchType(inputContent);
  const typeInfo = TYPE_MODULES[researchType];
  console.log(`✅ 调研类型：${researchType}型 — ${typeInfo.name}`);

  const client = new Anthropic({ apiKey: cfg.apiKey });
  const systemPrompt = buildSystemPrompt(researchType);

  // 创建输出目录
  if (!fs.existsSync(cfg.outputDir)) {
    fs.mkdirSync(cfg.outputDir, { recursive: true });
  }

  const now = new Date();
  const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
  const productTag = (contextSummary.product || "report").slice(0, 10).replace(/[/\s]/g, "_");

  const modulesOutput = {};
  const { modules } = typeInfo;
  const total = modules.length;

  console.log(`\n开始生成 ${total} 个模块...\n`);

  for (let i = 0; i < modules.length; i++) {
    const [moduleKey, moduleName] = modules[i];
    process.stdout.write(`[${i + 1}/${total}] 正在生成「${moduleName}」...`);

    const prompt = buildModulePrompt(moduleKey, inputContent);
    const result = await callClaude(client, prompt, systemPrompt, cfg);
    modulesOutput[moduleKey] = result;

    console.log(` 完成（${result.length}字）`);

    if (i < modules.length - 1) {
      await new Promise((r) => setTimeout(r, 800));
    }
  }

  console.log(`\n✅ 所有模块生成完毕\n`);

  const wordPath = path.join(cfg.outputDir, `市场调研报告_${productTag}_${dateStr}.docx`);
  await generateWordReport(
    modulesOutput, modules, researchType, contextSummary, wordPath
  );

  const pptxPath = path.join(cfg.outputDir, `市场调研报告_${productTag}_${dateStr}.pptx`);
  await generatePptxReport(
    modulesOutput, modules, researchType, contextSummary, pptxPath
  );

  const jsonPath = path.join(cfg.outputDir, `raw_output_${productTag}_${dateStr}.json`);
  fs.writeFileSync(
    jsonPath,
    JSON.stringify({
      researchType,
      typeName: typeInfo.name,
      contextSummary,
      modules: modulesOutput,
      generatedAt: now.toISOString(),
    }, null, 2),
    "utf-8"
  );

  console.log(`\n${"=".repeat(60)}`);
  console.log(`✅ 全部完成！输出文件：`);
  console.log(`   📄 Word：${wordPath}`);
  console.log(`   📊 PPT： ${pptxPath}`);
  console.log(`   🗂  JSON：${jsonPath}`);
  console.log(`${"=".repeat(60)}\n`);
}

runResearch(CONFIG);
