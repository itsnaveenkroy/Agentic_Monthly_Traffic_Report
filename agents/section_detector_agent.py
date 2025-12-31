"""
Section Detector Agent - LangGraph Node
Detects sections in the Excel sheet and extracts data into DataFrames.
"""

import pandas as pd
from typing import Dict, Any, List, Tuple
from openpyxl.worksheet.worksheet import Worksheet


class SectionDetectorAgent:
    """
    Agent responsible for detecting sections and extracting data.
    Identifies section headers and table boundaries.
    """
    
    def __init__(self):
        """Initialize the Section Detector Agent."""
        # No hard-coded headers - detect dynamically
    
    def detect_sections(self, worksheet: Worksheet) -> List[Dict[str, Any]]:
        """
        Detect ALL sections in the worksheet dynamically.
        A section is defined as:
        - A non-empty cell in column A (section name)
        - Followed by a row containing 'Month' in column B
        
        Args:
            worksheet: openpyxl worksheet object
            
        Returns:
            List of section metadata dictionaries
        """
        sections = []
        max_row = worksheet.max_row
        
        print(f"  [DEBUG] Scanning {max_row} rows for sections...")
        
        # Scan ALL rows to find section headers
        for row_idx in range(1, max_row + 1):
            row = worksheet[row_idx]
            
            # Check if column A has a value (potential section name)
            col_a_value = row[0].value
            if not col_a_value:
                continue
            
            # Skip if it's just 'Month' or other non-section markers
            col_a_str = str(col_a_value).strip()
            if col_a_str.lower() in ['month', 'total', '% change', '%change']:
                continue
            
            # Check if column B has 'Month' (this indicates header row)
            col_b_value = row[1].value if len(row) > 1 else None
            if col_b_value and str(col_b_value).strip().lower() == 'month':
                # Found a valid section!
                section_info = {
                    'name': col_a_str,
                    'header_row': row_idx,
                    'data_start_row': row_idx + 1,  # Data starts next row
                }
                sections.append(section_info)
                print(f"  [DEBUG] Found section '{col_a_str}' at row {row_idx}")
        
        # Calculate end rows for each section
        for i, section in enumerate(sections):
            if i < len(sections) - 1:
                # End row is just before next section
                section['data_end_row'] = sections[i + 1]['header_row'] - 1
            else:
                # Last section goes to end of sheet
                section['data_end_row'] = max_row
        
        return sections
    
    def extract_section_data(self, worksheet: Worksheet, section_info: Dict[str, Any]) -> pd.DataFrame:
        """
        Extract data for a specific section into a DataFrame.
        
        Args:
            worksheet: openpyxl worksheet object
            section_info: Section metadata dictionary
            
        Returns:
            pandas DataFrame with section data
        """
        # Headers are in the same row as section name, starting from column B (index 1)
        header_row_idx = section_info['header_row']
        headers = []
        
        # Read all cells in header row, skipping column A (which has section name)
        header_row = worksheet[header_row_idx]
        max_col = len(header_row)
        
        # Skip first column (index 0) - that's the section name
        for i, cell in enumerate(header_row[1:], start=1):  
            if cell.value:
                headers.append(str(cell.value).strip())
            else:
                headers.append(f"Column_{i+1}")
        
        print(f"  [DEBUG] Header row {header_row_idx} contents (skipping col A): {headers[:10]}")
        
        # Extract data rows (starting from row after headers, skipping column A)
        data_rows = []
        for row_idx in range(section_info['data_start_row'], section_info['data_end_row'] + 1):
            row = worksheet[row_idx]
            # Skip column A (index 0) - read from column B onwards
            row_data = [cell.value for cell in row[1:len(headers)+1]]
            
            # Stop if we hit an empty row
            if not any(row_data):
                break
            
            # Check if this is a meaningful data row (has at least the month column filled)
            if row_data and row_data[0]:  # First data column should have a value (Month)
                data_rows.append(row_data)
        
        # Create DataFrame
        df = pd.DataFrame(data_rows, columns=headers[:len(data_rows[0])] if data_rows else headers)
        
        print(f"  [DEBUG] Extracted {len(df)} rows for section with {len(df.columns)} columns")
        print(f"  [DEBUG] Column names: {list(df.columns)[:10]}...")  # First 10 columns
        if len(df) > 0:
            print(f"  [DEBUG] First row sample: {df.iloc[0].head(5).tolist()}")
        
        # FIX 2: Filter out ONLY % Change rows (keep Total for calculations)
        if len(df) > 0 and len(df.columns) > 0:
            month_col = 'Month' if 'Month' in df.columns else df.columns[0]
            # Keep month rows AND Total row, but remove % Change row
            df = df[~df[month_col].astype(str).str.lower().str.contains('% change|%change', na=False)]
            # Also remove completely empty rows
            df = df.dropna(how='all')
        
        # Reset index
        df = df.reset_index(drop=True)
        
        print(f"  [DEBUG] After filtering: {len(df)} valid data rows (including Total)")
        
        return df
    
    def execute(self, state: Dict[str, Any]) -> Dict[str, Any]:
        """
        Detect sections and prepare them for processing.
        
        Args:
            state: Current LangGraph state
            
        Returns:
            Updated state with sections list
        """
        print(f"\n{'='*60}")
        print("SECTION DETECTOR AGENT")
        print(f"{'='*60}")
        
        worksheet = state['worksheet']
        
        # Detect all sections
        sections = self.detect_sections(worksheet)
        
        print(f"[+] Detected {len(sections)} sections:")
        for i, section in enumerate(sections, 1):
            print(f"  {i}. {section['name']} (Row {section['header_row']})")
        
        # Store sections in state
        state['sections'] = sections
        state['current_section_index'] = 0
        
        print(f"[+] Section detection complete")
        print(f"{'='*60}\n")
        
        return state


def create_section_detector_node():
    """
    Factory function to create Section Detector node for LangGraph.
    
    Returns:
        Callable node function
    """
    agent = SectionDetectorAgent()
    return agent.execute
