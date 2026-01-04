# 🤖 GA4 Traffic Analysis - Agentic AI with LangGraph

An intelligent **Agentic AI system** powered by **LangGraph** that automates Google Analytics 4 (GA4) traffic report analysis. The system features multi-agent orchestration to calculate Year-over-Year (YoY) and Month-over-Month (MoM) metrics, generate AI-powered executive summaries, and produce professionally formatted Excel reports.

---

## 📋 Table of Contents

- [Overview](#-overview)
- [Key Features](#-key-features)
- [Architecture](#-architecture)
- [Project Structure](#-project-structure)
- [Installation](#-installation)
- [Configuration](#-configuration)
- [Usage](#-usage)
- [How It Works](#-how-it-works)
- [Input Data Format](#-input-data-format)
- [Calculation Rules](#-calculation-rules)
- [Output](#-output)
- [Agents Overview](#-agents-overview)
- [Troubleshooting](#-troubleshooting)
- [Technology Stack](#-technology-stack)
- [License](#-license)

---

## 🎯 Overview

This project demonstrates production-ready **Agentic AI** architecture using **LangGraph** for workflow orchestration. It processes GA4 traffic data through a series of specialized agents, each responsible for a specific task: data ingestion, section detection, metric calculation, AI-powered summary generation, and Excel output formatting.

**Use Case**: Automatically transform raw GA4 traffic reports into comprehensive analytical reports with AI-generated insights, saving hours of manual analysis and reporting work.

---

## ✨ Key Features

### Core Capabilities
✅ **LangGraph Multi-Agent Orchestration** - Five specialized agents working in coordinated workflow  
✅ **Intelligent Section Detection** - Automatically identifies traffic categories (Total Visits, Engaged Sessions, Referral, Paid, Social Media)  
✅ **Safe Metric Calculation** - Production-grade zero/null handling, never produces `NaN`, `Infinity`, or division-by-zero errors  
✅ **YOY % Calculation** - Year-over-Year growth analysis (2024 → 2025)  
✅ **LM % Calculation** - Month-over-Month trend analysis for 2025 data  

### AI-Powered Features
✅ **LLM-Generated Executive Summaries** - Professional, business-focused insights for each traffic section  
✅ **Context-Aware Prompting** - Different prompts for active vs. inactive sections  
✅ **Multi-LLM Support** - Compatible with OpenAI (GPT-4), OpenRouter, and Ollama (local)  

### Visual Formatting & Presentation
✅ **Professional Summary Headers** - Bold "Summary / Insights :" labels above each summary block  
✅ **Subtle Borders** - Light gray (#D3D3D3) borders around summary blocks for clear visual separation  
✅ **Intelligent Color Coding** - Automatic text color based on trend keywords:
   - 🟢 Green (#00B050) for "upward" trends (positive growth)
   - 🔵 Blue (#22577A) for "declined" trends (eye-pleasing, not aggressive)
   - ⚫ Black for neutral/mixed trends or no keywords  
✅ **Typography Excellence** - Century Gothic font, 12pt, with proper text wrapping and alignment  
✅ **Presentation-Only Formatting** - Visual enhancements without modifying logic or calculations  
✅ **Priority Logic** - "Upward" keyword takes precedence when both keywords exist (indicates overall positive trend)  

### Production Features
✅ **Excel Integration** - Direct read/write with proper formatting, cell merging, and alignment  
✅ **Dynamic Section Handling** - No hardcoded assumptions, adapts to any section count  
✅ **Comprehensive Error Handling** - Validation, fallbacks, and detailed logging  
✅ **Environment-Based Configuration** - `.env` file for API keys and paths  
✅ **Smart % Change Calculation** - Uses first year (2023) as reference baseline with Jan-Aug partial year comparison  
✅ **Total Row Auto-Calculation** - Automatically calculates sum of all month values (replaces formulas with actual values)  
✅ **Debug Logging** - Detailed console output for tracking metric writes and calculations  

---

## 🏗️ Architecture

### Multi-Agent Workflow

```
┌─────────────────────────────────────────────────────────────────┐
│                         USER INPUT                               │
│                      python main.py                              │
└──────────────────────────┬──────────────────────────────────────┘
                           │
                           ▼
               ┌──────────────────────┐
               │  LangGraph StateGraph │
               │    Orchestration      │
               └──────────────────────┘
                           │
        ┌──────────────────┼──────────────────┐
        │                  │                  │
        ▼                  ▼                  ▼
┌──────────────┐  ┌──────────────┐  ┌──────────────┐
│ Agent 1      │  │ Agent 2      │  │ Agent 3      │
│ Excel Reader │→ │ Section      │→ │ Metrics      │
│              │  │ Detector     │  │ Calculator   │
└──────────────┘  └──────────────┘  └──────┬───────┘
                                            │
                  ┌─────────────────────────┘
                  │
        ┌─────────┴─────────┐
        ▼                   ▼
┌──────────────┐   ┌──────────────┐
│ Agent 4      │   │ Agent 5      │
│ Summary      │→  │ Excel Writer │
│ Generator+LLM│   │              │
└──────────────┘   └──────────────┘
                          │
                          ▼
                   OUTPUT EXCEL
```

### Workflow Sequence

1. **ExcelReaderAgent** - Loads workbook and worksheet into shared state
2. **SectionDetectorAgent** - Scans worksheet to identify all traffic sections dynamically
3. **FOR EACH SECTION** (loop):
   - Extract section data into pandas DataFrame
   - **MetricsCalculatorAgent** - Calculate YOY and LM percentages with safety checks
   - **SummaryGeneratorAgent** - Generate executive summary via LLM
   - **ExcelWriterAgent** - Write metrics and summary back to Excel
4. **Save** - Final workbook saved to output path

---

## 📁 Project Structure

```
ga4_langgraph_agent/
│
├── main.py                           # Main entry point with LLM initialization
├── requirements.txt                  # Python dependencies
├── .env                              # Environment configuration (API keys, paths)
├── .gitignore                        # Git ignore rules
│
├── README.md                         # Comprehensive documentation
├── QUICKSTART.md                     # 3-minute setup guide
├── ARCHITECTURE.md                   # Detailed architecture documentation
├── PROJECT_COMPLETE.txt              # Project completion notes
│
├── agents/                           # LangGraph Agent Modules
│   ├── __init__.py
│   ├── excel_reader_agent.py         # Loads Excel workbook
│   ├── section_detector_agent.py     # Dynamically detects sections
│   ├── metrics_calculator_agent.py   # Calculates YOY & LM with safety
│   ├── summary_generator_agent.py    # LLM-powered summary generation
│   └── excel_writer_agent.py         # Writes results back to Excel
│
├── graph/                            # LangGraph Orchestration
│   ├── __init__.py
│   └── traffic_analysis_graph.py     # StateGraph workflow definition
│
├── utils/                            # Utility Functions
│   ├── __init__.py
│   ├── excel_utils.py                # Safe calculation functions
│   └── prompt_templates.py           # LLM prompt templates
│
├── data/                             # Data Files
│   ├── input_report.xlsx             # Input GA4 traffic data
│   └── output_report.xlsx            # Generated output
│
└── [verification scripts]            # Testing utilities
    ├── verify_all_sections.py
    ├── verify_output.py
    ├── verify_pct_change.py
    ├── verify_colors.py
    ├── check_lm_in_change_row.py
    └── debug_metrics.py
```

---

## 🚀 Installation

### Prerequisites

- **Python 3.9+** (3.10 or 3.11 recommended)
- **pip** (Python package manager)
- **Virtual environment** (recommended for dependency isolation)
- **API Key** (OpenAI, OpenRouter, or Ollama installed locally)

### Quick Setup (3 Steps)

#### Step 1: Create Virtual Environment

Navigate to the project directory and create an isolated Python environment:

**macOS / Linux:**
```bash
cd /path/to/ga4_langgraph_agent
python3 -m venv venv
source venv/bin/activate
```

**Windows:**
```bash
cd \path\to\ga4_langgraph_agent
python -m venv venv
venv\Scripts\activate
```

You should see `(venv)` prefix in your terminal.

#### Step 2: Install Dependencies

Install all required Python packages:

```bash
pip install -r requirements.txt
```

**Dependencies Installed:**
- `langgraph` - Agent workflow orchestration
- `langchain` & `langchain-openai` - LLM integration
- `python-dotenv` - Environment variable management
- `pandas` - Data processing
- `openpyxl` - Excel file handling
- `numpy` - Numerical operations

#### Step 3: Configure Environment Variables

Create or edit the `.env` file in the project root:

```bash
# Open in text editor
nano .env  # or: code .env, vim .env
```

**Configuration Template:**

```env
# ============================================
# LLM CONFIGURATION
# ============================================
LLM_PROVIDER=openai          # Options: openai | openrouter | ollama
OPENAI_API_KEY=sk-proj-xxxxx # Your OpenAI API key
OPENROUTER_API_KEY=sk-or-xxx # Your OpenRouter key (if using OpenRouter)
OLLAMA_MODEL=llama3          # Ollama model name (if using Ollama)

# ============================================
# FILE PATHS
# ============================================
INPUT_EXCEL_PATH=data/input_report.xlsx
OUTPUT_EXCEL_PATH=data/output_report.xlsx
```

⚠️ **Important:** Replace placeholder values with your actual API keys!

### Verify Installation

```bash
# Check Python version
python --version  # Should be 3.9+

# Verify packages installed
pip list | grep langgraph

# Test LLM connection (optional)
python -c "from langchain_openai import ChatOpenAI; print('✅ LangChain installed')"
```



---

## ⚙️ Configuration

### LLM Provider Options

The system supports three LLM providers with automatic initialization:

#### 1️⃣ OpenAI (Recommended for Production)

**Best For:** Production use, highest quality summaries  
**Model:** GPT-4  
**Cost:** ~$0.03 per 1K tokens (input), ~$0.06 per 1K tokens (output)

```env
LLM_PROVIDER=openai
OPENAI_API_KEY=sk-proj-your_actual_key_here
```

**Get API Key:** [platform.openai.com/api-keys](https://platform.openai.com/api-keys)

**Features:**
- Highest quality executive summaries
- Best contextual understanding
- Reliable, fast responses
- Production-grade stability

#### 2️⃣ OpenRouter (Cost-Effective Alternative)

**Best For:** Budget-conscious use, multiple model access  
**Model:** Mistral Devstral 2512 (free tier available)  
**Cost:** Free tier available, paid models from $0.001/1K tokens

```env
LLM_PROVIDER=openrouter
OPENROUTER_API_KEY=sk-or-your_actual_key_here
```

**Get API Key:** [openrouter.ai/keys](https://openrouter.ai/keys)

**Features:**
- Access to multiple LLM providers
- Free tier available
- Good quality summaries
- Flexible model selection

#### 3️⃣ Ollama (Local, Free)

**Best For:** Privacy-focused, no API costs, offline use  
**Model:** Llama 3 (or any Ollama model)  
**Cost:** Free (requires local compute)

```env
LLM_PROVIDER=ollama
OLLAMA_MODEL=llama3
```

**Setup:**
1. Install Ollama: [ollama.ai](https://ollama.ai/)
2. Pull model: `ollama pull llama3`
3. Start Ollama service: `ollama serve`

**Features:**
- 100% private, no data sent externally
- No API costs
- Runs offline
- Requires GPU for best performance

### File Path Configuration

Customize input/output paths in `.env`:

```env
# Use relative paths (from project root)
INPUT_EXCEL_PATH=data/input_report.xlsx
OUTPUT_EXCEL_PATH=data/output_report.xlsx

# Or absolute paths
INPUT_EXCEL_PATH=/Users/username/Documents/traffic_data.xlsx
OUTPUT_EXCEL_PATH=/Users/username/Documents/analysis_output.xlsx
```

---

## 📊 Usage

### Running the Analysis

Once configured, execute the main script:

```bash
# Ensure virtual environment is activated
source venv/bin/activate  # macOS/Linux
# or: venv\Scripts\activate  # Windows

# Run the analysis
python main.py
```

### Execution Flow

The system will automatically:

1. ✅ **Load Configuration** - Read `.env` file and validate settings
2. ✅ **Initialize LLM** - Connect to configured provider (OpenAI/OpenRouter/Ollama)
3. ✅ **Validate Environment** - Check file existence and directory structure
4. ✅ **Build LangGraph** - Initialize all 5 agents and workflow
5. ✅ **Execute Workflow**:
   - Load Excel workbook
   - Detect all traffic sections
   - For each section:
     - Extract data
     - Calculate YOY and LM percentages
     - Generate AI summary
     - Write results to Excel
6. ✅ **Save Output** - Write final report to `OUTPUT_EXCEL_PATH`
7. ✅ **Print Summary** - Display execution statistics

### Expected Output (Console)

```
============================================================
GA4 LANGGRAPH TRAFFIC ANALYSIS - AGENTIC AI SYSTEM
============================================================

[+] Loaded .env configuration
[+] Environment variables loaded
[+] Input file exists: data/input_report.xlsx
[+] Output directory exists: data

Initializing LLM Provider: OPENAI
[+] OpenAI GPT-4 initialized

============================================================
BUILDING LANGGRAPH WORKFLOW
============================================================

[+] LangGraph workflow built successfully
[+] Agents initialized:
  1. ExcelReaderAgent
  2. SectionDetectorAgent
  3. MetricsCalculatorAgent
  4. SummaryGeneratorAgent
  5. ExcelWriterAgent

[Processing sections...]

============================================================
EXECUTION SUMMARY
============================================================
[+] Sections processed: 5
[+] YOY percentages calculated
[+] LM percentages calculated
[+] Executive summaries generated
[+] Results written to Excel

============================================================
WORKFLOW COMPLETED SUCCESSFULLY
============================================================

Open the output file to view results:
   data/output_report.xlsx
```

### View Results

```bash
# macOS
open data/output_report.xlsx

# Linux
xdg-open data/output_report.xlsx

# Windows
start data/output_report.xlsx
```

---

## 🧠 How It Works

### LangGraph StateGraph Workflow

The system uses **LangGraph's StateGraph** to orchestrate a multi-agent workflow with shared state management.

#### State Management

All agents share a TypedDict state containing:

```python
class TrafficAnalysisState(TypedDict):
    workbook: Any              # openpyxl Workbook object
    worksheet: Any             # openpyxl Worksheet object
    sheet_name: str            # Active worksheet name
    sections: list             # Detected sections metadata
    current_section_index: int # Loop iteration counter
    section_name: str          # Current section being processed
    section_info: dict         # Section boundaries and metadata
    section_dataframe: Any     # pandas DataFrame with raw data
    calculated_metrics: Any    # DataFrame with YOY/LM columns
    summary_text: str          # LLM-generated summary
    all_summaries: list        # All summaries (tracking)
```

#### Workflow Graph

```
START
  │
  ├─► [ExcelReaderAgent]
  │    Load workbook & worksheet into state
  │
  ├─► [SectionDetectorAgent]
  │    Scan worksheet, detect all sections
  │    Initialize section_index = 0
  │
  └─► [LOOP: For each section]
       │
       ├─► [ProcessSectionNode]
       │    Extract current section data
       │    Create pandas DataFrame
       │
       ├─► [MetricsCalculatorAgent]
       │    Calculate YOY % (2024→2025)
       │    Calculate LM % (month-over-month)
       │    Handle zero/null values safely
       │
       ├─► [SummaryGeneratorAgent]
       │    Analyze calculated metrics
       │    Build LLM prompt with context
       │    Generate executive summary
       │
       ├─► [ExcelWriterAgent]
       │    Write YOY/LM percentages to Excel
       │    Write summary to merged cells
       │    Apply formatting
       │
       ├─► [IncrementIndexNode]
       │    current_section_index += 1
       │
       └─► [DECISION: More sections?]
            ├─ YES → Loop back to ProcessSectionNode
            └─ NO  → END (save workbook)
```

### Key Design Patterns

1. **Shared State** - All agents read/write to central state dictionary
2. **Sequential Processing** - Agents execute in defined order
3. **Loop Control** - Conditional edges manage section iteration
4. **Safe Calculations** - Zero/null checks prevent errors
5. **LLM Integration** - Context-aware prompts for quality summaries

---

## 📥 Input Data Format

### Excel Structure

Each section contains:

| Month | Sessions (GA4) Year-2023 | Sessions (GA4) Year-2024 | Sessions (GA4) Year-2025 | YOY % (2024-2025) | LM % (2025) |
|-------|--------------------------|--------------------------|--------------------------|-------------------|-------------|
| January | 15000 | 16500 | 18000 | **[Empty]** | **[Empty]** |
| February | 15200 | 16800 | 18500 | **[Empty]** | **[Empty]** |
| ... | ... | ... | ... | **[Empty]** | **[Empty]** |

### Supported Sections

- ✅ Total Visits (Sessions)
- ✅ Engaged Sessions
- ✅ Referral Traffic
- ✅ Paid Traffic
- ✅ Social Media Traffic
- ✅ Custom sections (auto-detected)

---

## 🧮 Calculation Rules

### YOY % (Year-over-Year)

**Formula:** `((2025 - 2024) / 2024) × 100`

**Conditions:**
- ✅ Only calculate if 2024 value exists AND > 0
- ✅ Only calculate if 2025 value exists
- ❌ Leave blank if 2024 = 0 or null
- ❌ Never write `0%`, `NaN`, or `Infinity`

### LM % (Last Month / Month-over-Month)

**Formula:** `((Current Month - Previous Month) / Previous Month) × 100`

**Conditions:**
- ✅ Only calculate if previous month exists AND > 0
- ✅ Only calculate if current month exists
- ❌ January LM % is **always blank** (no previous month)
- ❌ Leave blank if previous month = 0 or null

### % Change Row Calculation

**Important:** The % Change row uses **Year-2023 as the reference baseline** for all subsequent year comparisons.

**Column-wise behavior:**
- **Month column:** "% Change" label
- **Year-2023 (first numeric):** EMPTY (reference baseline - 0% change from itself)
- **Year-2024:** `(Total_2024 - Total_2023) / Total_2023 × 100`
- **Year-2025:** `(Jan-Aug_2025 - Jan-Aug_2023) / Jan-Aug_2023 × 100` + " (till Aug)"
- **YOY %:** `(Jan-Aug_2025 - Jan-Aug_2024) / Jan-Aug_2024 × 100` + " (till Aug)"
- **LM %:** EMPTY

**Jan-Aug Partial Year Comparison:**
- For 2025 (partial year), only Jan-Aug months are summed and compared to Jan-Aug of reference year
- Adds "(till Aug)" suffix to indicate partial year comparison
- Ensures fair comparison when full year data isn't available

### Total Row Calculation

**Behavior:**
- ✅ Automatically calculates sum of all month values (Jan through Dec)
- ✅ Writes actual numeric values (not formulas)
- ✅ Excludes the Total row itself from the sum
- ✅ Rounds to whole numbers (no decimals)
- ✅ Applies to all year columns (2023, 2024, 2025)
- ❌ Does not sum YOY % or LM % columns

### Special Handling

**Empty Sections** (all zeros):
- Skip YOY and LM calculations
- Generate summary: *"No measurable traffic was recorded..."*

**Total Rows / % Change Rows:**
- Automatically skipped from month-by-month metric calculations
- % Change row is calculated separately with its own logic
- Total row is calculated with actual sums of monthly data

---

## 📤 Output

### Generated Excel File

The output Excel file (`data/output_report.xlsx`) contains:

1. **Original Data** - Preserved from input
2. **YOY % Column** - Calculated percentages (e.g., `15.25%`)
3. **LM % Column** - Month-over-month percentages (e.g., `8.33%`)
4. **Executive Summaries** - Professional formatted summaries with:
   - **"Summary / Insights" Header** - Bold header above each summary block
   - **Light Gray Borders** - Subtle borders around summary blocks
   - **Color-Coded Text** - Automatic keyword-based formatting:
     - 🟢 **Green text (#00B050)** - Upward trends (positive growth)
     - 🔵 **Blue text (#22577A)** - Decline trends (negative growth)
     - ⚫ **Black text** - Neutral or mixed trends
   - **Century Gothic Font** - Professional 12pt font
   - **Text Wrapping** - Enabled for readability
   - **Top Alignment** - Clean presentation

### Visual Formatting Features

**Summary Block Presentation:**
- ✅ Header row with "Summary / Insights" label (bold)
- ✅ Light gray borders (thin, #D3D3D3) around summary
- ✅ Merged cells spanning the section height
- ✅ Automatic color coding based on trend keywords
- ✅ Professional typography (Century Gothic, 12pt)
- ✅ Text wrapping enabled for multi-line summaries

**Color Logic with Priority:**
- Summary contains "upward" → **Green text (#00B050)** - Takes PRIORITY (even if "declined" also present)
- Summary contains "declined" (without "upward") → **Blue text (#22577A)** - Secondary (eye-pleasing, not aggressive)
- No keywords or mixed signals → **Black text** - Neutral default

**Rationale:** The "upward" keyword indicates the overall positive trend direction, so it takes precedence when both positive and negative indicators are mentioned in the same summary.
- Summary contains neither → **Black text** (neutral)

### Summary Example

> **Total Visits (Sessions)** showed strong Year-over-Year growth from 2024 to 2025, with an average increase of 18.5% across all months. Month-over-Month performance in 2025 exhibited moderate volatility, ranging from -3.2% to +12.7%, indicating seasonal fluctuations. Overall, the upward trajectory demonstrates positive user engagement trends throughout the reporting period.

---

## 🤖 Agents Overview

### Agent Responsibilities

#### 1️⃣ ExcelReaderAgent

**Purpose:** Initialize Excel workbook access and load worksheet

**Inputs:**
- `input_path` (from .env configuration)

**Outputs:**
- `workbook` - openpyxl Workbook object
- `worksheet` - Active worksheet reference  
- `sheet_name` - Name of the worksheet

**Operations:**
- Validates file existence
- Loads workbook using openpyxl
- Stores workbook and worksheet in shared state

**Error Handling:**
- Raises `FileNotFoundError` if input file doesn't exist
- Validates workbook can be opened

---

#### 2️⃣ SectionDetectorAgent

**Purpose:** Dynamically identify all traffic sections in the worksheet

**Inputs:**
- `worksheet` - Excel worksheet object from state

**Outputs:**
- `sections` - List of section metadata dictionaries:
  ```python
  {
    'name': 'Total Visits (Sessions)',
    'header_row': 1,
    'data_start_row': 3,
    'data_end_row': 14
  }
  ```

**Detection Algorithm:**
1. Scan all rows in worksheet
2. Identify section headers (Column A has value, Column B has "Month")
3. Calculate section boundaries
4. Skip rows with "Total" or "% Change" in calculations

**Detected Sections (typical):**
- Total Visits (Sessions)
- Engaged Sessions
- Referral Traffic
- Paid Traffic
- Social Media Traffic
- *[Any custom sections matching the pattern]*

---

#### 3️⃣ MetricsCalculatorAgent

**Purpose:** Calculate YOY and LM percentages with production-grade safety

**Inputs:**
- `section_dataframe` - pandas DataFrame with raw section data
- `section_name` - Current section name

**Outputs:**
- `calculated_metrics` - DataFrame with added YOY and LM columns

**Calculation Logic:**

**YOY % (2024 → 2025):**
```python
IF 2024_value > 0 AND 2025_value exists:
    YOY = ((2025 - 2024) / 2024) × 100
ELSE:
    YOY = None  # Leave blank
```

**LM % (Month-over-Month 2025):**
```python
IF previous_month > 0 AND current_month exists:
    LM = ((current - previous) / previous) × 100
ELSE:
    LM = None  # Leave blank

SPECIAL: January LM = None (no previous month)
```

**Safety Features:**
- ✅ Never divides by zero
- ✅ Returns `None` instead of `NaN`/`Infinity`
- ✅ Handles empty cells gracefully
- ✅ Skips "Total" and "% Change" rows
- ✅ Validates numeric data types

---

#### 4️⃣ SummaryGeneratorAgent

**Purpose:** Generate executive summaries using LLM analysis

**Inputs:**
- `section_name` - Current section name
- `calculated_metrics` - DataFrame with YOY and LM data
- `llm_client` - Configured LLM instance

**Outputs:**
- `summary_text` - 3-5 sentence executive summary

**LLM Prompting Strategy:**

```
┌─────────────────────────────────────────┐
│  PROMPT STRUCTURE                       │
├─────────────────────────────────────────┤
│  1. Role Definition                     │
│     "You are an expert data analyst..." │
│                                         │
│  2. Context                             │
│     - Section name                      │
│     - Metrics summary (avg, range)      │
│                                         │
│  3. Task Requirements                   │
│     - Describe YOY trend                │
│     - Describe LM behavior              │
│     - Use business language             │
│                                         │
│  4. Constraints                         │
│     - 3-5 sentences                     │
│     - Professional tone                 │
│     - No technical jargon               │
└─────────────────────────────────────────┘
```

**Special Cases:**
- **Active sections** → Trend analysis with growth patterns
- **Inactive sections** → "No measurable traffic" summary
- **LLM failure** → Fallback template summary

---

#### 5️⃣ ExcelWriterAgent

**Purpose:** Write calculated metrics and summaries back to Excel

**Inputs:**
- `worksheet` - Excel worksheet object
- `calculated_metrics` - DataFrame with YOY/LM results
- `summary_text` - Generated summary
- `section_info` - Section boundaries and metadata

**Outputs:**
- Updated Excel worksheet with formatted results

**Operations:**
1. **Locate columns** - Find YOY (Column E) and LM (Column F)
2. **Write metrics** - Write percentage values row-by-row as formatted strings
3. **Format percentages** - Display as "15.25%" (not 0.1525)
4. **Write summary** - Insert into Column H with merged cells
5. **Apply styling** - Text wrapping, alignment, cell merging

**Formatting Features:**
- ✅ Merged cells for executive summary (spans multiple rows)
- ✅ Text wrapping enabled for readability
- ✅ Top alignment for summary text
- ✅ Percentage formatting for metrics
- ✅ Preserves original data
- ✅ **"Summary / Insights" header** - Bold label above each summary
- ✅ **Light gray borders** - Subtle borders (#D3D3D3) around summaries
- ✅ **Color-coded text** based on trend keywords:
  - **Green text (#00B050)** for upward/positive trends
  - **Blue text (#22577A)** for decline/negative trends (eye-pleasing color)
  - **Black text** for neutral or mixed trends
- ✅ **Century Gothic font, 12pt** - Professional typography
- ✅ **Presentation-only formatting** - No changes to logic or calculations

---

## 🛠️ Technology Stack

| Component | Technology | Version | Purpose |
|-----------|-----------|---------|---------|
| **Orchestration** | LangGraph | Latest | Multi-agent workflow management |
| **AI Framework** | LangChain | Latest | Unified LLM interface |
| **LLM Provider** | OpenAI GPT-4 | gpt-4 | Executive summary generation |
| | OpenRouter | Mistral Devstral | Alternative LLM access |
| | Ollama | Llama 3 | Local LLM execution |
| **Excel I/O** | openpyxl | 3.x | Read/write Excel files |
| **Data Processing** | pandas | 2.x | DataFrame operations |
| **Calculations** | NumPy | 1.x | Safe numeric operations |
| **Configuration** | python-dotenv | 1.x | Environment variables |
| **Language** | Python | 3.9+ | Core programming language |

### Key Dependencies

```
langgraph          # Agent workflow orchestration
langchain          # LLM framework
langchain-openai   # OpenAI integration
python-dotenv      # .env file loading
pandas             # Data manipulation
openpyxl           # Excel file handling
numpy              # Numerical computations
```

---

## 🐛 Troubleshooting

### Common Issues and Solutions

#### Error: "OPENAI_API_KEY not set"

**Cause:** API key not configured in `.env` file or contains placeholder value

**Solution:**
```bash
# Edit .env file
nano .env

# Replace this line:
OPENAI_API_KEY=your_key_here

# With your actual key:
OPENAI_API_KEY=sk-proj-xxxxxxxxxxxxx
```

---

#### Error: "Input Excel file not found"

**Cause:** File path in `.env` is incorrect or file doesn't exist

**Solution:**
1. Verify file exists: `ls data/input_report.xlsx`
2. Check `.env` path matches actual location
3. Use absolute path if needed:
   ```env
   INPUT_EXCEL_PATH=/full/path/to/input_report.xlsx
   ```

---

#### Error: "Module 'langgraph' not found"

**Cause:** Dependencies not installed or virtual environment not activated

**Solution:**
```bash
# Activate virtual environment first
source venv/bin/activate  # macOS/Linux
venv\Scripts\activate     # Windows

# Install dependencies
pip install -r requirements.txt

# Verify installation
pip list | grep langgraph
```

---

#### Error: "LLM Not Responding" or Timeout

**Cause:** API key invalid, network issues, or rate limiting

**Solution:**
1. **Verify API key:**
   ```python
   # Test OpenAI connection
   python -c "from openai import OpenAI; client = OpenAI(); print('✅ Connected')"
   ```

2. **Check internet connection**

3. **Try alternative provider:**
   ```env
   # Switch to Ollama (local, no network needed)
   LLM_PROVIDER=ollama
   OLLAMA_MODEL=llama3
   ```

4. **Check rate limits** - OpenAI free tier has request limits

---

#### Empty or Generic Summaries

**Cause:** LLM provider configuration issue or insufficient credits

**Solution:**
1. Verify API key has active credits/quota
2. Check LLM provider status page
3. Try different LLM model:
   ```env
   # For OpenRouter
   LLM_PROVIDER=openrouter
   OPENROUTER_API_KEY=your_key_here
   ```

---

#### Incorrect Percentage Calculations

**Cause:** Input data format mismatch or incorrect column detection

**Solution:**
1. **Verify column headers** match expected format:
   - Column B: "Month"
   - Column C: Contains "Year-2023" or "2023"
   - Column D: Contains "Year-2024" or "2024"
   - Column E: Contains "Year-2025" or "2025"

2. **Check data types** - Ensure numeric values, not text

3. **Review Total/% Change rows** - System automatically skips these

---

#### Permission Denied when Saving Output

**Cause:** Output file is open in Excel or insufficient write permissions

**Solution:**
1. Close `output_report.xlsx` if open in Excel
2. Check file permissions: `ls -la data/output_report.xlsx`
3. Change output path to writable location:
   ```env
   OUTPUT_EXCEL_PATH=/tmp/output_report.xlsx
   ```

---

### Debug Mode

Enable verbose logging for troubleshooting:

```python
# Add to main.py temporarily
import logging
logging.basicConfig(level=logging.DEBUG)
```

---

### Getting Help

If issues persist:

1. ✅ Verify all prerequisites installed (Python 3.9+, packages)
2. ✅ Check `.env` configuration matches documentation
3. ✅ Ensure input Excel format matches expected structure
4. ✅ Review console output for specific error messages
5. ✅ Test with sample data provided in `data/` directory

---

## 📚 Additional Resources

- **LangGraph Documentation:** [langchain-ai.github.io/langgraph](https://langchain-ai.github.io/langgraph/)
- **LangChain Docs:** [python.langchain.com](https://python.langchain.com)
- **OpenAI API Reference:** [platform.openai.com/docs](https://platform.openai.com/docs)
- **Architecture Details:** See [ARCHITECTURE.md](ARCHITECTURE.md)
- **Quick Start:** See [QUICKSTART.md](QUICKSTART.md)


---

## 🙏 Acknowledgments

- **LangGraph** - Agentic workflow orchestration
- **LangChain** - LLM integration framework
- **OpenPyXL** - Excel file manipulation
- **Pandas** - Data processing

---

## 📧 Support

For issues or questions:
1. Check the [Troubleshooting](#-troubleshooting) section
2. Review `.env` configuration
3. Ensure all dependencies are installed
4. Verify input Excel file format matches expected structure

---

**Built with ❤️ using LangGraph and Agentic AI principles**
