"""
Summary Generator Agent - LangGraph Node
Generates executive summaries using LLM analysis.
"""

import pandas as pd
from typing import Dict, Any, Optional
import sys
import os

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.prompt_templates import (
    EXECUTIVE_SUMMARY_PROMPT,
    EMPTY_SECTION_SUMMARY_PROMPT,
    build_data_summary,
    build_empty_section_context
)
from utils.excel_utils import is_section_empty


class SummaryGeneratorAgent:
    """
    Agent responsible for generating executive summaries using LLM.
    """
    
    def __init__(self, llm_client):
        """
        Initialize the Summary Generator Agent.
        
        Args:
            llm_client: LangChain LLM client (ChatOpenAI, ChatOllama, etc.)
        """
        self.llm = llm_client
    
    def identify_metric_columns(self, df: pd.DataFrame) -> Dict[str, str]:
        """
        Identify YOY and LM columns in DataFrame.
        
        Args:
            df: DataFrame with calculated metrics
            
        Returns:
            Dictionary with 'yoy' and 'lm' column names
        """
        metric_cols = {'yoy': None, 'lm': None}
        
        for col in df.columns:
            col_str = str(col).lower()
            if 'yoy' in col_str and ('2024' in col_str or '2025' in col_str):
                metric_cols['yoy'] = col
            elif 'lm' in col_str and '2025' in col_str:
                metric_cols['lm'] = col
        
        return metric_cols
    
    def identify_year_columns(self, df: pd.DataFrame) -> Dict[str, str]:
        """
        Identify which columns contain data for each year.
        
        Args:
            df: DataFrame with section data
            
        Returns:
            Dictionary mapping years to column names
        """
        year_columns = {}
        
        for col in df.columns:
            col_str = str(col).lower()
            if '2024' in col_str or 'year-2024' in col_str:
                year_columns['2024'] = col
            elif '2025' in col_str or 'year-2025' in col_str:
                year_columns['2025'] = col
        
        return year_columns
    
    def generate_summary(self, section_name: str, df: pd.DataFrame) -> str:
        """
        FIX 7: Generate executive summary based on CALCULATED metrics, not raw data.
        
        Args:
            section_name: Name of the section
            df: DataFrame with calculated metrics
            
        Returns:
            Generated summary text
        """
        # FIX 7: Check if we have any calculated metrics (YOY or LM)
        metric_cols = self.identify_metric_columns(df)
        
        # Section is inactive ONLY if NO calculated metrics exist
        has_yoy = metric_cols['yoy'] and metric_cols['yoy'] in df.columns and df[metric_cols['yoy']].notna().any()
        has_lm = metric_cols['lm'] and metric_cols['lm'] in df.columns and df[metric_cols['lm']].notna().any()
        
        if not has_yoy and not has_lm:
            # Generate empty section summary
            print(f"  [DEBUG] No calculated metrics found - generating inactive summary")
            prompt = EMPTY_SECTION_SUMMARY_PROMPT.format(
                section_name=section_name
            )
        else:
            # Generate regular executive summary
            print(f"  [DEBUG] Found metrics: YOY={has_yoy}, LM={has_lm} - generating trend summary")
            
            # Build data summary
            if metric_cols['yoy'] and metric_cols['lm']:
                data_summary = build_data_summary(df, metric_cols['yoy'], metric_cols['lm'])
            else:
                data_summary = "Limited metric data available for analysis."
            
            prompt = EXECUTIVE_SUMMARY_PROMPT.format(
                section_name=section_name,
                data_summary=data_summary
            )
        
        # Call LLM to generate summary
        try:
            response = self.llm.invoke(prompt)
            
            # Extract text from response
            if hasattr(response, 'content'):
                summary = response.content.strip()
            else:
                summary = str(response).strip()
            
            return summary
        
        except Exception as e:
            print(f"  [!] Error generating summary with LLM: {e}")
            # FIX 7: Fallback based on calculated metrics, not raw data
            metric_cols = self.identify_metric_columns(df)
            has_yoy = metric_cols['yoy'] and metric_cols['yoy'] in df.columns and df[metric_cols['yoy']].notna().any()
            has_lm = metric_cols['lm'] and metric_cols['lm'] in df.columns and df[metric_cols['lm']].notna().any()
            
            if not has_yoy and not has_lm:
                return f"No measurable traffic was recorded in {section_name} during the analyzed period."
            else:
                return f"Analysis of {section_name} metrics shows varying patterns across the reporting period. Further investigation recommended."
    
    def execute(self, state: Dict[str, Any]) -> Dict[str, Any]:
        """
        Generate executive summary for current section.
        
        Args:
            state: Current LangGraph state
            
        Returns:
            Updated state with summary text
        """
        print(f"\n{'='*60}")
        print("SUMMARY GENERATOR AGENT")
        print(f"{'='*60}")
        
        section_name = state['section_name']
        df = state['calculated_metrics']
        
        print(f"[+] Generating summary for: {section_name}")
        
        # Generate summary using LLM
        summary = self.generate_summary(section_name, df)
        
        print(f"[+] Summary generated ({len(summary)} characters)")
        print(f"\nGenerated Summary:")
        print(f"{'─'*60}")
        print(summary)
        print(f"{'─'*60}")
        
        # Update state
        state['summary_text'] = summary
        
        print(f"\n[+] Summary generation complete")
        print(f"{'='*60}\n")
        
        return state


def create_summary_generator_node(llm_client):
    """
    Factory function to create Summary Generator node for LangGraph.
    
    Args:
        llm_client: LangChain LLM client
        
    Returns:
        Callable node function
    """
    agent = SummaryGeneratorAgent(llm_client)
    return agent.execute
