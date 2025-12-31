"""
Metrics Calculator Agent - LangGraph Node
Calculates YOY and LM percentages with safe zero/null handling.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Optional
import sys
import os

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.excel_utils import (
    calculate_yoy_percentage,
    calculate_lm_percentage,
    is_skip_row,
    is_section_empty
)


class MetricsCalculatorAgent:
    """
    Agent responsible for calculating YOY and LM percentages.
    Implements safe calculation logic with zero/null handling.
    """
    
    def __init__(self):
        """Initialize the Metrics Calculator Agent."""
        pass
    
    def identify_year_columns(self, df: pd.DataFrame) -> Dict[str, str]:
        """
        Identify which columns contain data for each year.
        
        Args:
            df: DataFrame with section data
            
        Returns:
            Dictionary mapping years to column names
        """
        year_columns = {}
        
        print(f"  [DEBUG] Available columns: {list(df.columns)}")
        
        for col in df.columns:
            col_str = str(col).lower()
            # Match patterns like "Sessions (GA4) Year-2023" or "year-2023" or "2023"
            if 'year-2023' in col_str or 'year 2023' in col_str or col_str.endswith('2023'):
                year_columns['2023'] = col
                print(f"  [DEBUG] Found 2023 column: {col}")
            elif 'year-2024' in col_str or 'year 2024' in col_str or col_str.endswith('2024'):
                year_columns['2024'] = col
                print(f"  [DEBUG] Found 2024 column: {col}")
            elif 'year-2025' in col_str or 'year 2025' in col_str or col_str.endswith('2025'):
                year_columns['2025'] = col
                print(f"  [DEBUG] Found 2025 column: {col}")
        
        print(f"  [DEBUG] Year columns mapped: {year_columns}")
        return year_columns
    
    def calculate_metrics(self, df: pd.DataFrame, section_name: str) -> pd.DataFrame:
        """
        Calculate YOY and LM percentages for the DataFrame.
        
        Args:
            df: DataFrame with section data
            section_name: Name of the section
            
        Returns:
            DataFrame with calculated metrics
        """
        # Make a copy to avoid modifying original
        result_df = df.copy()
        
        # Identify year columns
        year_cols = self.identify_year_columns(result_df)
        
        # Debug: Check if section has actual data
        print(f"  [DEBUG] Section data preview:")
        if '2024' in year_cols:
            col_2024 = year_cols['2024']
            sample_values = result_df[col_2024].head(3).tolist()
            print(f"  [DEBUG] Sample 2024 values: {sample_values}")
        if '2025' in year_cols:
            col_2025 = year_cols['2025']
            sample_values = result_df[col_2025].head(3).tolist()
            print(f"  [DEBUG] Sample 2025 values: {sample_values}")
        
        # Check if section is empty
        if is_section_empty(result_df, list(year_cols.values())):
            print(f"  [!] Section '{section_name}' has no measurable data (all zeros or empty)")
            result_df['YOY %'] = None
            result_df['LM %'] = None
            return result_df
        
        # Get month column (usually first column)
        month_col = result_df.columns[0]
        
        # Initialize YOY and LM columns
        yoy_col = 'YOY % (2024-2025)'
        lm_col = 'LM % (2025)'
        
        # Find existing YOY and LM columns or create new ones
        for col in result_df.columns:
            col_str = str(col).lower()
            if 'yoy' in col_str and '2024' in col_str and '2025' in col_str:
                yoy_col = col
            elif 'lm' in col_str and '2025' in col_str:
                lm_col = col
        
        # Create columns if they don't exist
        if yoy_col not in result_df.columns:
            result_df[yoy_col] = None
        if lm_col not in result_df.columns:
            result_df[lm_col] = None
        
        # FIX 4 & 6: Calculate YOY percentages row-by-row and handle % Change row
        if '2024' in year_cols and '2025' in year_cols:
            col_2024 = year_cols['2024']
            col_2025 = year_cols['2025']
            
            for idx, row in result_df.iterrows():
                # Skip if this is a % change row
                if is_skip_row(row):
                    continue
                
                # Check if this is Total row - skip YOY for Total (will be in % Change row instead)
                month_val = row.iloc[0] if len(row) > 0 else None
                is_total_row = pd.notna(month_val) and 'total' in str(month_val).lower()
                
                if is_total_row:
                    result_df.at[idx, yoy_col] = None
                    print(f"  [DEBUG] Skipping YOY for Total row (goes in % Change row instead)")
                    continue
                
                val_2024 = row[col_2024]
                val_2025 = row[col_2025]
                
                # Convert to numeric
                try:
                    val_2024 = pd.to_numeric(val_2024, errors='coerce')
                    val_2025 = pd.to_numeric(val_2025, errors='coerce')
                except:
                    val_2024 = None
                    val_2025 = None
                
                # Calculate YOY
                yoy_result = calculate_yoy_percentage(val_2024, val_2025)
                result_df.at[idx, yoy_col] = yoy_result
        
        # Calculate LM percentages (Month-over-Month for 2025)
        if '2025' in year_cols:
            col_2025 = year_cols['2025']
            col_2024 = year_cols.get('2024')  # May be None
            
            # Track previous month value
            previous_value = None
            first_month = True
            december_2024_value = None
            
            # If 2024 column exists, find December 2024 value for January comparison
            if col_2024:
                for idx, row in result_df.iterrows():
                    if is_skip_row(row):
                        continue
                    month_val = row.iloc[0] if len(row) > 0 else None
                    if pd.notna(month_val) and 'december' in str(month_val).lower():
                        try:
                            december_2024_value = pd.to_numeric(row[col_2024], errors='coerce')
                            print(f"  [DEBUG] Found December 2024 value: {december_2024_value}")
                        except:
                            december_2024_value = None
                        break
            
            for idx, row in result_df.iterrows():
                # Skip if this is a % change row
                if is_skip_row(row):
                    continue
                
                # Check if this is Total row - skip LM for Total (doesn't make sense)
                month_val = row.iloc[0] if len(row) > 0 else None
                is_total_row = pd.notna(month_val) and 'total' in str(month_val).lower()
                
                if is_total_row:
                    result_df.at[idx, lm_col] = None
                    print(f"  [DEBUG] Skipping LM for Total row (not applicable)")
                    continue
                
                current_value = row[col_2025]
                
                # Convert to numeric
                try:
                    current_value = pd.to_numeric(current_value, errors='coerce')
                except:
                    current_value = None
                
                # First month (January) compares to December 2024 if available
                if first_month:
                    if december_2024_value is not None and pd.notna(december_2024_value):
                        lm_result = calculate_lm_percentage(current_value, december_2024_value)
                        result_df.at[idx, lm_col] = lm_result
                        print(f"  [DEBUG] January LM calculated: {lm_result}% (Jan 2025 vs Dec 2024)")
                    else:
                        result_df.at[idx, lm_col] = None
                        print(f"  [DEBUG] January LM skipped (no December 2024 data)")
                    previous_value = current_value
                    first_month = False
                    continue
                
                # Calculate LM if we have previous value
                if previous_value is not None:
                    lm_result = calculate_lm_percentage(current_value, previous_value)
                    result_df.at[idx, lm_col] = lm_result
                else:
                    result_df.at[idx, lm_col] = None
                
                # Update previous value for next iteration
                previous_value = current_value
        
        # FIX 6: Store totals for % Change calculation
        # Track if we need to calculate % Change row
        total_2024 = None
        total_2025 = None
        if '2024' in year_cols and '2025' in year_cols:
            col_2024 = year_cols['2024']
            col_2025 = year_cols['2025']
            # Sum non-null numeric values
            total_2024 = pd.to_numeric(result_df[col_2024], errors='coerce').sum()
            total_2025 = pd.to_numeric(result_df[col_2025], errors='coerce').sum()
        
        # Store total values in state for later use by writer
        result_df.attrs['total_2024'] = total_2024
        result_df.attrs['total_2025'] = total_2025
        
        print(f"  [DEBUG] Totals: 2024={total_2024}, 2025={total_2025}")
        
        return result_df
    
    def execute(self, state: Dict[str, Any]) -> Dict[str, Any]:
        """
        Calculate metrics for current section.
        
        Args:
            state: Current LangGraph state
            
        Returns:
            Updated state with calculated metrics
        """
        print(f"\n{'='*60}")
        print("METRICS CALCULATOR AGENT")
        print(f"{'='*60}")
        
        section_name = state['section_name']
        df = state['section_dataframe']
        
        print(f"[+] Processing section: {section_name}")
        print(f"[+] Data shape: {df.shape[0]} rows x {df.shape[1]} columns")
        
        # Calculate metrics
        result_df = self.calculate_metrics(df, section_name)
        
        # Count calculated values
        yoy_col = [col for col in result_df.columns if 'yoy' in str(col).lower()]
        lm_col = [col for col in result_df.columns if 'lm' in str(col).lower()]
        
        if yoy_col:
            yoy_count = result_df[yoy_col[0]].notna().sum()
            print(f"[+] Calculated {yoy_count} YOY percentages")
        
        if lm_col:
            lm_count = result_df[lm_col[0]].notna().sum()
            print(f"[+] Calculated {lm_count} LM percentages")
        
        # Update state
        state['section_dataframe'] = result_df
        state['calculated_metrics'] = result_df
        
        print(f"[+] Metrics calculation complete")
        print(f"{'='*60}\n")
        
        return state


def create_metrics_calculator_node():
    """
    Factory function to create Metrics Calculator node for LangGraph.
    
    Returns:
        Callable node function
    """
    agent = MetricsCalculatorAgent()
    return agent.execute
