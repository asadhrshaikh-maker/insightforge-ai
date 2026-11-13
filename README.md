# 🚀 Business Copilot AI V3

Enterprise, AI-powered Business Intelligence platform. Upload **any** Excel or CSV
file and the platform automatically builds an executive dashboard, generates AI
insights, performs root-cause analysis, lets you chat with your data, and exports
board-ready **PDF / PowerPoint / Excel / CSV / Email** reports — zero configuration.

Built with **Streamlit · Pandas · NumPy · Plotly · OpenPyXL · python-pptx · Matplotlib**.
The PDF renderer is a self-contained pure-Python engine (no system dependencies).

---

## ✨ Features

### Phase 1 — Auto Dashboard
- **Dataset detection**: Sales, Survey, Finance, HR, Operations, Inventory, Marketing, General.
- **9 KPIs**: Total Records, Total Columns, Missing Data %, Duplicate Rows, Date Range,
  Top Category, Primary Metric, Data Quality Score, Growth % (period-over-period).
- **Executive layout**: KPI row → Executive Summary → Charts → Insights.
- **Executive summary** written in readable management language.

### Phase 2 — AI Insights Engine (`🤖 AI Insights`)
- **What Happened**: top contributor, largest category, highest/lowest value, trends, changes.
- **Risks**: data quality, concentration (HHI), missing values, outliers (IQR),
  statistical anomalies (rolling z-score), performance decline.
- **Opportunities**: growth segments, emerging categories, underutilized segments, momentum.
- **Recommendations**: prioritized P1/P2/P3 business actions.

### Phase 2B — Root Cause Analysis (`🔎 Root Cause Analysis`)
- One click decomposes the latest period change across **every** detected dimension
  (Date, Category, Region, Department, Project, Status, Product, Channel, Location…).
- Ranks drivers by contribution, separates volume vs. value effects, shows a
  reconciling **waterfall bridge**, and writes plain-English explanations.

### Phase 3 — Chat with Data (`💬 Ask Your Data`)
- Natural-language Q&A grounded in your actual numbers.
- **Engine priority**: OpenAI (if an API key is supplied) → Ollama → Built-in Analyst.
- **Ollama** model dropdown (default `qwen2.5-coder`; `llama3`, `qwen`, `mistral`,
  `deepseek-r1` + any installed models via *Refresh*).
- The **Built-in Analyst** works fully offline (no LLM needed) and can render charts.

### Phase 4 — Reports (`📄 Reports`)
- Management Report, Executive Summary, KPI Summary, Risk Report, Opportunity Report,
  Root Cause Report — each downloadable as Markdown.
- **Phase 4B — PDF**: branded title/logo area, summary, KPI grid, charts, insights,
  recommendations, driver tables, column profile.
- **Phase 4C — PowerPoint**: 8 slides (Title, Summary, KPIs, Charts, Insights,
  Recommendations, Risk Analysis, Next Steps).
- **Phase 4D — Email**: subject + overview + insights + risks + recommendations,
  ready to paste into Outlook/Gmail.

### UI / UX
Premium dark theme, gradient background, glassmorphism cards, soft shadows, rounded
corners, large typography. Brand palette: `#3B82F6 #8B5CF6 #10B981 #F59E0B #EF4444`
on `#050B18 / #071326 / #0B2344`.

### Charts (auto-selected)
Bar, Top-10 horizontal bar, Pie, Donut, Line, Area, Segment trend lines, Scatter,
Box plot, Density heatmap, Correlation heatmap, Treemap, Sunburst.

### Global Filters
Date range, category/region/department/status/project multiselects, numeric range
sliders, column text filter, and full-text search — applied across every tab.

### Exports (`📥 Export`)
CSV, multi-sheet Excel, Insights Report (Markdown), PDF, PowerPoint.

---

## 🚀 Quick start

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate
pip install -r requirements.txt
python -m streamlit run app_v3.py