"""
Prompt templates for LLM-based summary generation.
Provides structured prompts for executive summary creation.
"""

EXECUTIVE_SUMMARY_PROMPT = """
You are an expert data analyst creating executive summaries for GA4 traffic reports.

**Section**: {section_name}

**Data Analysis**:
{data_summary}

**Your Task**:
Generate a professional, 3-5 sentence executive summary that:

1. Describes the Year-over-Year (YOY) trend from 2024 to 2025
   - Is there growth, decline, or mixed performance?
   - What is the overall direction?

2. Describes the Month-over-Month (LM) behavior within 2025
   - Is the trend stable, volatile, increasing, or decreasing?
   - Are there notable month-to-month changes?

3. Uses business-friendly, executive-level language
   - Avoid technical jargon
   - Focus on insights and implications
   - Be concise and actionable

**Important Rules**:
- Do NOT mention missing data, errors, or calculation issues
- Do NOT use phrases like "insufficient data" or "unable to calculate"
- If data is limited, focus on what IS available
- Maintain a professional, confident tone
- Use past tense for historical data
- Keep the summary between 3-5 sentences

**Tone**: Professional, analytical, executive-focused

**Output**: Provide ONLY the summary text, no additional formatting or commentary.

Summary:
"""


EMPTY_SECTION_SUMMARY_PROMPT = """
You are an expert data analyst creating executive summaries for GA4 traffic reports.

**Section**: {section_name}

**Observation**: This section shows no measurable traffic or engagement during the analyzed period.

**Your Task**:
Generate a professional, 2-3 sentence executive summary that:

1. States that no measurable activity was recorded
2. Suggests this may indicate:
   - Inactive channel
   - Data collection issues
   - Channel not utilized during this period
3. Recommends monitoring or investigation if appropriate

**Tone**: Professional, neutral, non-alarming

**Output**: Provide ONLY the summary text, no additional formatting or commentary.

Summary:
"""


TREND_ANALYSIS_PROMPT = """
Analyze the following metrics and provide a brief trend description:

**Metrics**:
{metrics_data}

**Task**:
Describe the trend in ONE sentence using these categories:
- "Strong growth" (YOY > 20%)
- "Moderate growth" (YOY 5-20%)
- "Stable" (YOY -5% to 5%)
- "Moderate decline" (YOY -20% to -5%)
- "Significant decline" (YOY < -20%)
- "Mixed performance" (inconsistent patterns)

Output only the trend description:
"""


def build_data_summary(df, yoy_col: str, lm_col: str) -> str:
    """
    Build a structured data summary for the prompt.
    
    Args:
        df: DataFrame with calculated metrics
        yoy_col: Column name for YOY percentages
        lm_col: Column name for LM percentages
        
    Returns:
        Formatted string summarizing the data
    """
    summary_lines = []
    
    # YOY Analysis
    yoy_values = df[yoy_col].dropna()
    if len(yoy_values) > 0:
        avg_yoy = yoy_values.mean()
        max_yoy = yoy_values.max()
        min_yoy = yoy_values.min()
        summary_lines.append(f"YOY Performance (2024→2025):")
        summary_lines.append(f"  - Average: {avg_yoy:.2f}%")
        summary_lines.append(f"  - Range: {min_yoy:.2f}% to {max_yoy:.2f}%")
        summary_lines.append(f"  - Months analyzed: {len(yoy_values)}")
    else:
        summary_lines.append("YOY Performance: Limited data for 2024-2025 comparison")
    
    summary_lines.append("")
    
    # LM Analysis
    lm_values = df[lm_col].dropna()
    if len(lm_values) > 0:
        avg_lm = lm_values.mean()
        max_lm = lm_values.max()
        min_lm = lm_values.min()
        summary_lines.append(f"Month-over-Month Performance (2025):")
        summary_lines.append(f"  - Average: {avg_lm:.2f}%")
        summary_lines.append(f"  - Range: {min_lm:.2f}% to {max_lm:.2f}%")
        summary_lines.append(f"  - Months with data: {len(lm_values)}")
        
        # Volatility indicator
        if len(lm_values) > 1:
            std_dev = lm_values.std()
            if std_dev > 20:
                summary_lines.append(f"  - Volatility: High (σ={std_dev:.2f})")
            elif std_dev > 10:
                summary_lines.append(f"  - Volatility: Moderate (σ={std_dev:.2f})")
            else:
                summary_lines.append(f"  - Volatility: Low (σ={std_dev:.2f})")
    else:
        summary_lines.append("Month-over-Month Performance: Insufficient month-to-month data")
    
    return "\n".join(summary_lines)


def build_empty_section_context(section_name: str) -> str:
    """
    Build context for empty/inactive sections.
    
    Args:
        section_name: Name of the section
        
    Returns:
        Formatted context string
    """
    return f"The '{section_name}' section contains no measurable traffic data for the analyzed period."
