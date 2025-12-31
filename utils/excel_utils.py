"""
Excel utility functions for reading and writing GA4 traffic data.
Handles safe data extraction and formatting.
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from typing import Dict, List, Tuple, Optional


def is_skip_row(row_data: pd.Series) -> bool:
    """
    Determine if a row should be skipped from calculations (ONLY % Change and empty rows).
    Total rows are KEPT for proper totals calculation.
    
    Args:
        row_data: pandas Series containing row data
        
    Returns:
        True if row should be skipped, False otherwise
    """
    if pd.isna(row_data.iloc[0]):  # Empty first cell
        return True
    
    first_cell = str(row_data.iloc[0]).strip().lower()
    
    # Skip ONLY % change rows (keep Total rows)
    if '% change' in first_cell or '%change' in first_cell:
        return True
    
    return False


def safe_percentage_calculation(numerator: float, denominator: float) -> Optional[float]:
    """
    Safely calculate percentage change with zero and null handling.
    
    Args:
        numerator: The difference value (current - previous)
        denominator: The base value (previous value)
        
    Returns:
        Percentage as float or None if calculation is invalid
    """
    # Check if denominator is zero, NaN, or None
    if pd.isna(denominator) or denominator == 0:
        return None
    
    # Check if numerator is valid
    if pd.isna(numerator):
        return None
    
    # Calculate percentage
    percentage = (numerator / denominator) * 100
    
    # Return None if result is infinite or NaN
    if np.isinf(percentage) or np.isnan(percentage):
        return None
    
    return round(percentage, 2)


def calculate_yoy_percentage(value_2024: float, value_2025: float) -> Optional[float]:
    """
    Calculate Year-over-Year percentage (2024 → 2025).
    
    Rules:
    - Only calculate if 2024 value exists AND > 0
    - Only calculate if 2025 value exists
    - Return None if invalid
    
    Args:
        value_2024: The 2024 value
        value_2025: The 2025 value
        
    Returns:
        YOY percentage or None
    """
    # Check if both values exist
    if pd.isna(value_2024) or pd.isna(value_2025):
        return None
    
    # Check if 2024 value is greater than 0
    if value_2024 <= 0:
        return None
    
    # Calculate YOY
    difference = value_2025 - value_2024
    return safe_percentage_calculation(difference, value_2024)


def calculate_lm_percentage(current_value: float, previous_value: float) -> Optional[float]:
    """
    Calculate Last Month percentage (Month-over-Month in 2025).
    
    Rules:
    - Only calculate if previous month value exists AND > 0
    - Only calculate if current month value exists
    - January LM % compares to December of previous year (2024)
    - Return None if invalid
    
    Args:
        current_value: Current month value
        previous_value: Previous month value
        
    Returns:
        LM percentage or None
    """
    # Check if both values exist
    if pd.isna(current_value) or pd.isna(previous_value):
        return None
    
    # Check if previous value is greater than 0
    if previous_value <= 0:
        return None
    
    # Calculate LM
    difference = current_value - previous_value
    return safe_percentage_calculation(difference, previous_value)


def is_section_empty(df: pd.DataFrame, year_columns: List[str]) -> bool:
    """
    FIX 3: Check if section is truly empty.
    A section is empty ONLY IF there is NO row where:
    - 2024 value exists and > 0
    - AND 2025 value exists
    
    Args:
        df: DataFrame containing section data
        year_columns: List of column names for year data
        
    Returns:
        True if section is empty/inactive, False otherwise
    """
    # Find 2024 and 2025 columns
    col_2024 = None
    col_2025 = None
    
    for col in year_columns:
        col_str = str(col).lower()
        if '2024' in col_str:
            col_2024 = col
        elif '2025' in col_str:
            col_2025 = col
    
    # If we don't have both columns, section is empty
    if not col_2024 or not col_2025:
        return True
    
    if col_2024 not in df.columns or col_2025 not in df.columns:
        return True
    
    # Check for valid rows: 2024 > 0 AND 2025 exists
    # Convert to numeric first to handle any string values
    col_2024_numeric = pd.to_numeric(df[col_2024], errors='coerce')
    col_2025_numeric = pd.to_numeric(df[col_2025], errors='coerce')
    
    valid_rows = df[
        col_2024_numeric.notna() &
        col_2025_numeric.notna() &
        (col_2024_numeric > 0)
    ]
    
    section_is_empty = len(valid_rows) == 0
    return section_is_empty


def detect_month_order(df: pd.DataFrame, month_column: str) -> List[str]:
    """
    Detect the chronological order of months in the data.
    
    Args:
        df: DataFrame containing month data
        month_column: Name of the column containing month names
        
    Returns:
        List of month names in chronological order
    """
    month_order = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ]
    
    # Extract unique months from DataFrame
    months_in_data = []
    for month in df[month_column]:
        if pd.notna(month) and str(month).strip():
            month_str = str(month).strip()
            if month_str in month_order:
                if month_str not in months_in_data:
                    months_in_data.append(month_str)
    
    # Sort by month order
    sorted_months = sorted(months_in_data, key=lambda x: month_order.index(x))
    
    return sorted_months


def write_summary_to_excel(worksheet, start_row: int, end_row: int, 
                           summary_column: str, summary_text: str):
    """
    Write summary text to Excel worksheet in specified column and row range.
    Merge cells vertically and set alignment.
    
    Args:
        worksheet: openpyxl worksheet object
        start_row: Starting row number (1-indexed)
        end_row: Ending row number (1-indexed)
        summary_column: Column letter (e.g., 'H')
        summary_text: The summary text to write
    """
    # Write text to first cell
    cell_address = f"{summary_column}{start_row}"
    worksheet[cell_address] = summary_text
    
    # Merge cells vertically
    merge_range = f"{summary_column}{start_row}:{summary_column}{end_row}"
    worksheet.merge_cells(merge_range)
    
    # Set alignment: wrap text, vertical top, horizontal left
    cell = worksheet[cell_address]
    cell.alignment = Alignment(
        wrap_text=True,
        vertical='top',
        horizontal='left'
    )


def format_percentage_for_excel(value: Optional[float]) -> str:
    """
    Format percentage value for Excel display.
    
    Args:
        value: Percentage value or None
        
    Returns:
        Formatted string for Excel (empty string if None)
    """
    if value is None:
        return ""
    return f"{value:.2f}%"
