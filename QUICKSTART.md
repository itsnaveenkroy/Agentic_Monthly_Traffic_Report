# 🚀 Quick Start Guide - GA4 LangGraph Agent

Get up and running with the GA4 Traffic Analysis system in **under 5 minutes**!

---

## ⚡ 3-Minute Setup

### Prerequisites

- Python 3.9+ installed
- Terminal/Command Prompt access
- OpenAI API key (or Ollama installed locally)

---

### Step 1: Environment Setup (1 minute)

```bash
# Navigate to project directory
cd /path/to/ga4_langgraph_agent

# Create virtual environment
python3 -m venv venv  # macOS/Linux
# or: python -m venv venv  # Windows

# Activate virtual environment
source venv/bin/activate  # macOS/Linux
# or: venv\Scripts\activate  # Windows

# Install dependencies
pip install -r requirements.txt
```

**Expected output:**
```
Successfully installed langgraph-X.X.X langchain-X.X.X ...
```

---

### Step 2: Configure API Key (30 seconds)

Create or edit `.env` file in project root:

```bash
# Quick edit (choose one)
nano .env        # Simple text editor
code .env        # VS Code
vim .env         # Vim
```

**Minimum configuration:**

```env
LLM_PROVIDER=openai
OPENAI_API_KEY=sk-proj-your_actual_key_here

INPUT_EXCEL_PATH=data/input_report.xlsx
OUTPUT_EXCEL_PATH=data/output_report.xlsx
```

⚠️ **Important:** Replace `sk-proj-your_actual_key_here` with your real OpenAI API key!

**Get API Key:** [platform.openai.com/api-keys](https://platform.openai.com/api-keys)

---

### Step 3: Run Analysis (30 seconds)

```bash
# Make sure virtual environment is activated (you should see (venv) in prompt)
python main.py
```

**What happens:**
1. ✅ Loads configuration
2. ✅ Initializes GPT-4
3. ✅ Reads `data/input_report.xlsx`
4. ✅ Detects traffic sections
5. ✅ Calculates YOY & LM percentages
6. ✅ Generates AI summaries
7. ✅ Saves to `data/output_report.xlsx`

---

### Step 4: View Results (10 seconds)

```bash
# Open output file
open data/output_report.xlsx        # macOS
xdg-open data/output_report.xlsx    # Linux
start data/output_report.xlsx       # Windows
```

**You'll see:**
- ✅ **Column E:** YOY % calculated (e.g., "15.25%")
- ✅ **Column F:** LM % calculated (e.g., "8.33%")
- ✅ **Column H:** AI-generated executive summaries

---

## 🎯 Using Your Own Data

### Option 1: Replace Input File

1. Prepare your Excel with this structure:
   - **Column A:** Section names
   - **Column B:** "Month" header, then Jan-Dec, Total, % Change
   - **Column C-E:** Year-2023, Year-2024, Year-2025 data
   - **Columns F-G:** Leave empty (will be filled)

2. Save as `data/input_report.xlsx`

3. Run: `python main.py`

---

### Option 2: Custom File Path

Edit `.env`:

```env
INPUT_EXCEL_PATH=/Users/yourname/Documents/my_traffic_data.xlsx
OUTPUT_EXCEL_PATH=/Users/yourname/Documents/my_analysis_output.xlsx
```

---

## 🔥 Common Commands Cheat Sheet

```bash
# Activate environment
source venv/bin/activate              # macOS/Linux
venv\Scripts\activate                 # Windows

# Run analysis
python main.py

# Deactivate environment when done
deactivate

# Check installed packages
pip list | grep langgraph

# Update dependencies
pip install --upgrade -r requirements.txt

# View logs during execution
python main.py 2>&1 | tee execution.log
```

---

## ⚙️ Alternative LLM Providers

Don't have OpenAI API key? No problem!

### Option A: Ollama (Free, Local, Private)

**Setup (5 minutes):**

1. **Install Ollama:**
   - Visit: [ollama.ai](https://ollama.ai/)
   - Download and install for your OS

2. **Pull a model:**
   ```bash
   ollama pull llama3
   ```

3. **Start Ollama service:**
   ```bash
   ollama serve
   ```

4. **Configure `.env`:**
   ```env
   LLM_PROVIDER=ollama
   OLLAMA_MODEL=llama3
   ```

5. **Run:**
   ```bash
   python main.py
   ```

**Benefits:**
- ✅ 100% free
- ✅ Runs offline
- ✅ Private (no data sent externally)
- ✅ No API costs

**Requirements:**
- ~8GB RAM minimum
- GPU recommended for speed

---

### Option B: OpenRouter (Free Tier Available)

**Setup (2 minutes):**

1. **Get API key:**
   - Visit: [openrouter.ai/keys](https://openrouter.ai/keys)
   - Sign up and copy key

2. **Configure `.env`:**
   ```env
   LLM_PROVIDER=openrouter
   OPENROUTER_API_KEY=sk-or-your_key_here
   ```

3. **Run:**
   ```bash
   python main.py
   ```

**Benefits:**
- ✅ Free tier available
- ✅ Access to multiple models
- ✅ Pay-per-use pricing
- ✅ No subscription needed

---

## 📊 What Each Agent Does

| Agent | Purpose | What It Does |
|-------|---------|--------------|
| **ExcelReaderAgent** | Data Loading | Opens workbook and loads worksheet into memory |
| **SectionDetectorAgent** | Section Discovery | Scans worksheet to find all traffic sections (Total Visits, Engaged Sessions, etc.) |
| **MetricsCalculatorAgent** | Calculations | Computes YOY % and LM % with zero-handling safety |
| **SummaryGeneratorAgent** | AI Insights | Calls LLM (GPT-4) to generate executive summaries |
| **ExcelWriterAgent** | Output Formatting | Writes metrics and summaries back to Excel with proper formatting |

**Workflow:** 
```
Load Excel → Detect Sections → [For Each Section: Calculate → Summarize → Write] → Save
```

---

## 📋 Expected Input Format

Your Excel file should have sections like this:

```
┌─────────────────────────────┬──────────┬──────────┬──────────┬──────────┬──────────┬──────────┐
│ Total Visits (Sessions)     │ Month    │ Year-2023│ Year-2024│ Year-2025│ YOY %    │ LM %     │
├─────────────────────────────┼──────────┼──────────┼──────────┼──────────┼──────────┼──────────┤
│                             │ January  │ 15,000   │ 16,500   │ 18,000   │ [Empty]  │ [Empty]  │
│                             │ February │ 15,200   │ 16,800   │ 18,500   │ [Empty]  │ [Empty]  │
│                             │ ...      │ ...      │ ...      │ ...      │ ...      │ ...      │
└─────────────────────────────┴──────────┴──────────┴──────────┴──────────┴──────────┴──────────┘
```

**Key Rules:**
- Section name in Column A (first row of section)
- "Month" in Column B (header row)
- Year columns: C=2023, D=2024, E=2025
- Columns F & G: Leave empty (will be calculated)

---

## ✅ Success Checklist

Before running, verify:

- [ ] Virtual environment created and activated
- [ ] Dependencies installed (`pip list | grep langgraph` shows results)
- [ ] `.env` file configured with API key
- [ ] `data/input_report.xlsx` exists or path configured
- [ ] LLM provider accessible (OpenAI/OpenRouter online, or Ollama running)

After running, check:

- [ ] Console shows "WORKFLOW COMPLETED SUCCESSFULLY"
- [ ] No error messages in output
- [ ] `data/output_report.xlsx` created
- [ ] Output file has YOY % (Column E) filled
- [ ] Output file has LM % (Column F) filled
- [ ] Output file has summaries (Column H) filled

---

## 🆘 Quick Troubleshooting

**"API key not set"**
```bash
# Edit .env and add your key
nano .env
```

**"Module not found"**
```bash
# Activate environment and reinstall
source venv/bin/activate
pip install -r requirements.txt
```

**"File not found"**
```bash
# Check file exists
ls data/input_report.xlsx

# Or use absolute path in .env
INPUT_EXCEL_PATH=/full/path/to/file.xlsx
```

**"LLM timeout"**
```bash
# Try local Ollama instead
ollama pull llama3
ollama serve
# Then update .env: LLM_PROVIDER=ollama
```

---

## 🚀 Next Steps

1. **Read Full Documentation:** [README.md](README.md)
2. **Learn Architecture:** [ARCHITECTURE.md](ARCHITECTURE.md)
3. **Customize Prompts:** Edit `utils/prompt_templates.py`
4. **Add Custom Metrics:** Extend `agents/metrics_calculator_agent.py`
5. **Change LLM Models:** Update LLM configuration in `main.py`

---

## 📚 Helpful Resources

- **LangGraph Tutorial:** [langchain-ai.github.io/langgraph](https://langchain-ai.github.io/langgraph/)
- **OpenAI API Docs:** [platform.openai.com/docs](https://platform.openai.com/docs)
- **Ollama Models:** [ollama.ai/library](https://ollama.ai/library)
- **OpenRouter Pricing:** [openrouter.ai/pricing](https://openrouter.ai/pricing)

---

**🎉 You're all set! Happy analyzing!**

```bash
python main.py
```
