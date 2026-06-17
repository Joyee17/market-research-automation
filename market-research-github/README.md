# Market Research Automation Tool

> McKinsey-standard automated market research report generator powered by Claude API. Judges research type, searches authoritative data sources, and exports executive-ready Word/PPT reports.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Python-3.9%2B-blue)](https://www.python.org/)
[![Node.js](https://img.shields.io/badge/Node.js-18%2B-green)](https://nodejs.org/)

---

## What This Does

Most market research reports fail at the executive level not because of insufficient data, but because of **vague conclusions, unsourced claims, and unactionable recommendations**. This tool enforces a strict analytical framework — modeled on top-tier consulting deliverables — across six research types, and automatically:

1. **Detects research type** from a structured input (Market Entry / Competitive Intelligence / Customer Insight / Market Penetration / Product & Pricing / Brand & Communication)
2. **Searches authoritative data sources** (World Bank, IEA, IRENA, BloombergNEF, McKinsey Global Institute, etc.) via Claude's web search tool
3. **Generates module-by-module analysis** enforcing a "What → Why → So What → Now What" insight structure
4. **Exports professional Word (.docx) and PowerPoint (.pptx) reports** ready for executive review

## Why It's Different

Most AI-generated business reports produce confident-sounding but unsupported text. This tool enforces hard quality gates at the prompt level:

- **No vague conclusions allowed** — "it depends" / "each has pros and cons" are explicitly banned outputs
- **Source-tiered data citation** — every data point is tagged by authority level (L1 international institutions → L4 unverified, flagged `[needs verification]`)
- **Scenario-based competitive analysis** — instead of static spec tables, comparisons are framed as "under [specific condition], competitor X fails at [quantified gap], we deliver [quantified advantage], translating to [customer financial value]"
- **Every action item requires a deadline, owner, and measurable success criterion**

## Six Research Frameworks Included

| Type | Core Question | Key Output |
|------|---------------|------------|
| A — Market Entry | Which market should we enter, and how? | Prioritized market matrix + GTM plan |
| B — Competitive Intelligence | What's the competitive threat, and how do we counter it? | Threat rating + sales battlecard |
| C — Customer & Demand Insight | What do customers actually need, and what will they pay for? | Demand priority matrix + updated ICP |
| D — Market Penetration | How do we grow share in an existing market? | ROI-ranked growth levers |
| E — Product & Pricing | What should we build next, and how should we price it? | Roadmap priority + pricing structure |
| F — Brand & Communication | How is our brand perceived, and is our messaging working? | Brand gap analysis + channel ROI |

## Quick Start

### Option 1: Python

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY="your-api-key"
python market_research_tool.py
```

### Option 2: Node.js

```bash
npm install
export ANTHROPIC_API_KEY="your-api-key"
node market_research_tool.mjs
```

### Before running

Fill in `00_internal_data_template.md` with your product info, target market(s), customer hypothesis, and company constraints. The tool reads this file as its primary input, then enriches it with live web search results before generating the report.

## Output

Running the tool produces three files in `output_reports/`:

```
市场调研报告_[product]_[timestamp].docx   ← Executive Word report
市场调研报告_[product]_[timestamp].pptx   ← Presentation deck
raw_output_[product]_[timestamp].json     ← Raw module outputs (for reuse/debugging)
```

## Architecture

```
Internal data template (user input)
        ↓
Research type auto-detection (6-way classifier)
        ↓
Module-by-module generation (Claude API + web_search tool)
        ↓
Quality gates enforced via system prompt
        ↓
Word (.docx) + PowerPoint (.pptx) export
```

## Data Source Authority Tiers

The tool enforces source citation discipline across four tiers:

- **L1** (highest): World Bank, IMF, IEA, IRENA, McKinsey Global Institute
- **L2** (high): BloombergNEF, Gartner, Reuters, Bloomberg, Wood Mackenzie
- **L3** (medium, cross-validation required): public company annual reports, trade press
- **L4** (reference only, must be flagged): Statista, LinkedIn, Crunchbase

Data older than 24 months or single-sourced market share figures are rejected by the prompt-level rules.

## Also Available

- `claude_system_prompt.md` — drop into a Claude.ai Project for a no-API, chat-based version of this workflow
- `coze_config.md` — full configuration for deploying this as a shareable team bot on Coze

## License

MIT — see [LICENSE](LICENSE)

## Tech Stack

`Claude API (Anthropic)` · `Python / Node.js` · `python-docx / docx (npm)` · `python-pptx / pptxgenjs`
