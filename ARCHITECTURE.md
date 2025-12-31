# 🏗️ Architecture & Design Documentation

## System Architecture

### High-Level Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                         USER                                     │
│                           │                                      │
│                           ▼                                      │
│                      python main.py                              │
└─────────────────────────────────────────────────────────────────┘
                             │
                             │ Loads .env
                             │ Initializes LLM
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│                   LANGGRAPH ORCHESTRATION                        │
│                  (traffic_analysis_graph.py)                     │
│                                                                  │
│  ┌────────────────────────────────────────────────────────┐    │
│  │              StateGraph Workflow                       │    │
│  │                                                        │    │
│  │  State Variables:                                     │    │
│  │  • workbook, worksheet                                │    │
│  │  • sections, current_section_index                    │    │
│  │  • section_name, section_info                         │    │
│  │  • section_dataframe, calculated_metrics              │    │
│  │  • summary_text, all_summaries                        │    │
│  └────────────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────────────┘
                             │
                             │
        ┌────────────────────┴────────────────────┐
        │                                         │
        ▼                                         ▼
┌──────────────┐                         ┌──────────────┐
│   AGENTS     │                         │   UTILITIES  │
└──────────────┘                         └──────────────┘
        │                                         │
        ├─ ExcelReaderAgent                      ├─ excel_utils.py
        ├─ SectionDetectorAgent                  │  └─ Safe calculations
        ├─ MetricsCalculatorAgent                │     Zero/null handling
        ├─ SummaryGeneratorAgent                 │
        └─ ExcelWriterAgent                      └─ prompt_templates.py
                                                     └─ LLM prompts
```

---

## LangGraph Workflow Flow

### Sequential Processing

```
START
  │
  ├─► [1] ExcelReaderAgent
  │    └─► Load workbook & worksheet
  │         Store in state
  │
  ├─► [2] SectionDetectorAgent  
  │    └─► Detect all sections
  │         Extract section metadata
  │         Initialize section loop
  │
  └─► [LOOP: For each section]
       │
       ├─► [3] ProcessSectionNode
       │    └─► Extract current section data
       │         Create DataFrame
       │
       ├─► [4] MetricsCalculatorAgent
       │    └─► Calculate YOY %
       │         Calculate LM %
       │         Handle zeros/nulls
       │
       ├─► [5] SummaryGeneratorAgent
       │    └─► Analyze metrics
       │         Call LLM
       │         Generate executive summary
       │
       ├─► [6] ExcelWriterAgent
       │    └─► Write YOY/LM to Excel
       │         Write summary to Excel
       │         Format cells
       │
       ├─► [7] IncrementIndexNode
       │    └─► Move to next section
       │
       └─► [DECISION] More sections?
            ├─ YES → Loop back to [3]
            └─ NO  → END
```

---

## Agent Responsibilities

### 1️⃣ ExcelReaderAgent

**Purpose:** Initialize Excel workbook access

**Inputs:**
- Input file path (from .env)

**Outputs:**
- `workbook`: openpyxl Workbook object
- `worksheet`: Active worksheet
- `sheet_name`: Worksheet name

**Operations:**
- Validates file exists
- Loads workbook with openpyxl
- Stores references in shared state

---

### 2️⃣ SectionDetectorAgent

**Purpose:** Identify all traffic sections in the Excel sheet

**Inputs:**
- `worksheet`: Excel worksheet object

**Outputs:**
- `sections`: List of section metadata
  ```python
  {
    'name': 'Total Visits (Sessions)',
    'header_row': 1,
    'data_start_row': 3,
    'data_end_row': 14
  }
  ```

**Operations:**
- Scans worksheet for section headers
- Identifies table boundaries
- Filters out "Total" and "% Change" rows
- Creates section metadata dictionary

**Detected Sections:**
- Total Visits (Sessions)
- Engaged Sessions
- Referral Traffic
- Paid Traffic
- Social Media Traffic

---

### 3️⃣ MetricsCalculatorAgent

**Purpose:** Calculate YOY and LM percentages safely

**Inputs:**
- `section_dataframe`: pandas DataFrame with raw data
- `section_name`: Current section name

**Outputs:**
- `calculated_metrics`: DataFrame with YOY % and LM % columns

**Calculation Logic:**

#### YOY % (2024 → 2025)
```python
IF 2024_value > 0 AND 2025_value exists:
    YOY = ((2025 - 2024) / 2024) × 100
ELSE:
    YOY = None  # Leave blank
```

#### LM % (Month-over-Month 2025)
```python
IF previous_month > 0 AND current_month exists:
    LM = ((current - previous) / previous) × 100
ELSE:
    LM = None  # Leave blank

SPECIAL CASE: January LM compares to December 2024 (cross-year)
```

**Safety Features:**
- Never divides by zero
- Returns `None` instead of NaN/Infinity
- Handles empty cells gracefully
- Skips "Total" and "% Change" rows

---

### 4️⃣ SummaryGeneratorAgent

**Purpose:** Generate executive summaries using LLM

**Inputs:**
- `section_name`: Section name
- `calculated_metrics`: DataFrame with YOY and LM

**Outputs:**
- `summary_text`: 3-5 sentence executive summary

**LLM Prompting Strategy:**

```
┌─────────────────────────────────────────┐
│  PROMPT STRUCTURE                       │
├─────────────────────────────────────────┤
│                                         │
│  1. Role Definition                     │
│     "You are an expert data analyst..." │
│                                         │
│  2. Context                             │
│     Section name                        │
│     Metrics summary (avg, range, etc.)  │
│                                         │
│  3. Task Requirements                   │
│     - Describe YOY trend                │
│     - Describe LM behavior              │
│     - Use business language             │
│                                         │
│  4. Constraints                         │
│     - 3-5 sentences                     │
│     - No technical jargon               │
│     - Professional tone                 │
│                                         │
└─────────────────────────────────────────┘
```

**Special Handling:**
- **Empty sections** → Dedicated prompt for zero-data scenarios
- **Limited data** → Focuses on available information
- **Fallback** → Uses template summary if LLM fails

---

### 5️⃣ ExcelWriterAgent

**Purpose:** Write results back to Excel

**Inputs:**
- `worksheet`: Excel worksheet
- `calculated_metrics`: DataFrame with results
- `summary_text`: Generated summary
- `section_info`: Section boundaries

**Outputs:**
- Updated Excel worksheet with:
  - YOY percentages in Column E
  - LM percentages in Column F
  - Executive summary in Column H (merged cells)

**Operations:**
1. Locate YOY/LM columns
2. Write percentage values row-by-row
3. Format as "15.25%" strings
4. Write summary to side panel
5. Merge cells vertically for summary
6. Apply text wrapping and alignment

---

## Data Flow Diagram

```
INPUT EXCEL
    │
    │ [ExcelReaderAgent]
    ▼
WORKBOOK OBJECT ──────────────┐
    │                         │
    │ [SectionDetectorAgent]  │
    ▼                         │
SECTIONS LIST                 │
    │                         │
    │ FOR EACH SECTION:       │
    ├─► SECTION DATAFRAME     │
    │        │                │
    │        │ [MetricsCalculatorAgent]
    │        ▼                │
    │   CALCULATED METRICS    │
    │        │                │
    │        │ [SummaryGeneratorAgent + LLM]
    │        ▼                │
    │   SUMMARY TEXT          │
    │        │                │
    │        │ [ExcelWriterAgent]
    │        ▼                │
    └───► UPDATED WORKSHEET ◄─┘
              │
              ▼
        OUTPUT EXCEL
```

---

## State Management

### Shared State Dictionary

The LangGraph workflow uses a `TypedDict` for state management:

```python
class TrafficAnalysisState(TypedDict):
    # Excel objects
    workbook: Any              # openpyxl Workbook
    worksheet: Any             # openpyxl Worksheet
    sheet_name: str            # Sheet name
    
    # Section tracking
    sections: list             # All detected sections
    current_section_index: int # Current loop position
    section_name: str          # Current section name
    section_info: dict         # Section metadata
    
    # Data processing
    section_dataframe: Any     # pandas DataFrame
    calculated_metrics: Any    # DataFrame with YOY/LM
    
    # Results
    summary_text: str          # Generated summary
    all_summaries: list        # All summaries (tracking)
```

### State Transitions

```
Initial State:
└─► workbook=None, sections=[], current_section_index=0

After ExcelReaderAgent:
└─► workbook=<Workbook>, worksheet=<Worksheet>

After SectionDetectorAgent:
└─► sections=[{...}, {...}], current_section_index=0

Per Section Loop:
└─► section_name="Total Visits"
    section_dataframe=<DataFrame>
    calculated_metrics=<DataFrame with YOY/LM>
    summary_text="..."
    current_section_index += 1
```

---

## Error Handling Strategy

### Input Validation

```python
# File existence
if not os.path.exists(input_path):
    raise FileNotFoundError(...)

# API key validation
if not api_key or api_key == 'your_key_here':
    raise ValueError("API key not configured")
```

### Safe Calculations

```python
# Zero division prevention
if denominator == 0 or pd.isna(denominator):
    return None  # Not 0%, not NaN

# Infinity check
if np.isinf(result) or np.isnan(result):
    return None
```

### LLM Fallbacks

```python
try:
    summary = llm.invoke(prompt)
except Exception as e:
    # Use template summary
    summary = generate_fallback_summary(section_name)
```

---

## Performance Considerations

### Optimization Points

1. **Single Excel Load** - Workbook loaded once, reused for all sections
2. **Batch Processing** - All sections processed in one workflow run
3. **Efficient DataFrame Operations** - Vectorized pandas calculations
4. **Minimal LLM Calls** - One summary per section (not per row)

### Scalability

- ✅ Handles multiple sections automatically
- ✅ No hardcoded row numbers
- ✅ Supports variable-length sections
- ✅ Memory-efficient state management

---

## Extension Points

### Adding New Agents

1. Create new agent file in `agents/`
2. Implement `execute(state) -> state` method
3. Add node to graph in `traffic_analysis_graph.py`
4. Update state schema if needed

### Custom Calculations

1. Add calculation function to `utils/excel_utils.py`
2. Call from `MetricsCalculatorAgent`
3. Update DataFrame with new column

### Custom Prompts

1. Add prompt template to `utils/prompt_templates.py`
2. Use in `SummaryGeneratorAgent`
3. Customize for your use case

---

## Technology Stack

| Component | Technology | Purpose |
|-----------|-----------|---------|
| **Orchestration** | LangGraph | Agent workflow management |
| **LLM Integration** | LangChain | Unified LLM interface |
| **LLM Provider** | OpenAI GPT-4 | Summary generation |
| **Excel I/O** | openpyxl | Read/write Excel files |
| **Data Processing** | pandas | DataFrame operations |
| **Calculations** | NumPy | Safe numeric operations |
| **Config** | python-dotenv | Environment variables |

---

## Design Principles

1. **Modularity** - Each agent has single responsibility
2. **Safety First** - Extensive zero/null handling
3. **Stateful Workflow** - Shared state across agents
4. **LLM-Augmented** - AI generates human-readable insights
5. **Production-Ready** - Error handling, logging, validation
6. **Configurable** - .env-based configuration
7. **Extensible** - Easy to add new agents/features

---

**Built with Agentic AI Architecture Principles** 🏗️
