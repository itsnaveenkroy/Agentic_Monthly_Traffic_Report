"""
Traffic Analysis Graph - LangGraph Orchestration
Orchestrates the flow of agents for GA4 traffic analysis.
"""

from typing import Dict, Any, TypedDict, Annotated
from langgraph.graph import StateGraph, END
import sys
import os

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from agents.excel_reader_agent import create_excel_reader_node
from agents.section_detector_agent import create_section_detector_node
from agents.metrics_calculator_agent import create_metrics_calculator_node
from agents.summary_generator_agent import create_summary_generator_node
from agents.excel_writer_agent import create_excel_writer_node


class TrafficAnalysisState(TypedDict):
    """
    Shared state for the LangGraph workflow.
    
    State variables:
    - workbook: openpyxl Workbook object
    - worksheet: openpyxl Worksheet object
    - sheet_name: Name of the active worksheet
    - sections: List of detected section metadata
    - current_section_index: Index of current section being processed
    - section_name: Name of current section
    - section_info: Metadata for current section
    - section_dataframe: pandas DataFrame with section data
    - calculated_metrics: DataFrame with YOY and LM calculations
    - summary_text: Generated executive summary
    - all_summaries: List of all generated summaries (for tracking)
    """
    workbook: Any
    worksheet: Any
    sheet_name: str
    sections: list
    current_section_index: int
    section_name: str
    section_info: dict
    section_dataframe: Any
    calculated_metrics: Any
    summary_text: str
    all_summaries: list


def process_section_node(state: Dict[str, Any]) -> Dict[str, Any]:
    """
    Node that extracts current section data from worksheet.
    
    Args:
        state: Current LangGraph state
        
    Returns:
        Updated state with section data
    """
    print(f"\n{'='*60}")
    print("PROCESSING SECTION NODE")
    print(f"{'='*60}")
    
    # Get current section
    sections = state['sections']
    current_idx = state['current_section_index']
    
    if current_idx >= len(sections):
        print("[+] All sections processed")
        print(f"{'='*60}\n")
        return state
    
    section_info = sections[current_idx]
    section_name = section_info['name']
    
    print(f"[+] Processing section {current_idx + 1}/{len(sections)}: {section_name}")
    
    # Import section detector to extract data
    from agents.section_detector_agent import SectionDetectorAgent
    detector = SectionDetectorAgent()
    
    # Extract section data
    section_df = detector.extract_section_data(state['worksheet'], section_info)
    
    print(f"[+] Extracted {len(section_df)} rows of data")
    
    # Update state
    state['section_name'] = section_name
    state['section_info'] = section_info
    state['section_info']['input_path'] = state.get('input_path')  # Pass input path for % Change
    state['section_dataframe'] = section_df
    
    print(f"{'='*60}\n")
    
    return state


def should_continue_sections(state: Dict[str, Any]) -> str:
    """
    Conditional edge to determine if more sections need processing.
    
    Args:
        state: Current LangGraph state
        
    Returns:
        "continue" to process next section, "end" to finish
    """
    sections = state['sections']
    current_idx = state['current_section_index']
    
    if current_idx < len(sections):
        return "process_next"
    else:
        return "end"


def increment_section_index(state: Dict[str, Any]) -> Dict[str, Any]:
    """
    Increment the section index to move to next section.
    
    Args:
        state: Current LangGraph state
        
    Returns:
        Updated state with incremented index
    """
    state['current_section_index'] += 1
    return state


class TrafficAnalysisGraph:
    """
    LangGraph workflow for GA4 traffic analysis.
    """
    
    def __init__(self, input_excel_path: str, llm_client):
        """
        Initialize the Traffic Analysis Graph.
        
        Args:
            input_excel_path: Path to input Excel file
            llm_client: LangChain LLM client for summary generation
        """
        self.input_path = input_excel_path
        self.llm_client = llm_client
        self.graph = None
        
    def build_graph(self) -> StateGraph:
        """
        Build the LangGraph workflow.
        
        Returns:
            Compiled StateGraph
        """
        # Create state graph
        workflow = StateGraph(TrafficAnalysisState)
        
        # Create agent nodes
        excel_reader = create_excel_reader_node(self.input_path)
        section_detector = create_section_detector_node()
        metrics_calculator = create_metrics_calculator_node()
        summary_generator = create_summary_generator_node(self.llm_client)
        excel_writer = create_excel_writer_node()
        
        # Add nodes to graph
        workflow.add_node("excel_reader", excel_reader)
        workflow.add_node("section_detector", section_detector)
        workflow.add_node("process_section", process_section_node)
        workflow.add_node("metrics_calculator", metrics_calculator)
        workflow.add_node("summary_generator", summary_generator)
        workflow.add_node("excel_writer", excel_writer)
        workflow.add_node("increment_index", increment_section_index)
        
        # Define workflow edges
        workflow.set_entry_point("excel_reader")
        
        # Linear flow for initialization
        workflow.add_edge("excel_reader", "section_detector")
        
        # Start section processing loop
        workflow.add_edge("section_detector", "process_section")
        
        # Section processing flow
        workflow.add_edge("process_section", "metrics_calculator")
        workflow.add_edge("metrics_calculator", "summary_generator")
        workflow.add_edge("summary_generator", "excel_writer")
        workflow.add_edge("excel_writer", "increment_index")
        
        # Conditional edge: continue to next section or end
        workflow.add_conditional_edges(
            "increment_index",
            should_continue_sections,
            {
                "process_next": "process_section",
                "end": END
            }
        )
        
        # Compile the graph
        self.graph = workflow.compile()
        
        return self.graph
    
    def run(self, recursion_limit: int = 100) -> Dict[str, Any]:
        """
        Execute the traffic analysis workflow.
        
        Args:
            recursion_limit: Maximum recursion depth for graph execution
        
        Returns:
            Final state after all sections processed
        """
        if not self.graph:
            self.build_graph()
        
        # Initialize state
        initial_state = {
            'workbook': None,
            'worksheet': None,
            'sheet_name': '',
            'sections': [],
            'current_section_index': 0,
            'section_name': '',
            'section_info': {},
            'section_dataframe': None,
            'calculated_metrics': None,
            'summary_text': '',
            'all_summaries': []
        }
        
        print("\n" + "="*60)
        print("STARTING LANGGRAPH TRAFFIC ANALYSIS WORKFLOW")
        print("="*60 + "\n")
        
        # Execute graph with recursion limit
        final_state = self.graph.invoke(
            initial_state,
            {"recursion_limit": recursion_limit}
        )
        
        print("\n" + "="*60)
        print("LANGGRAPH WORKFLOW COMPLETED")
        print("="*60 + "\n")
        
        return final_state


def create_traffic_analysis_graph(input_excel_path: str, llm_client) -> TrafficAnalysisGraph:
    """
    Factory function to create Traffic Analysis Graph.
    
    Args:
        input_excel_path: Path to input Excel file
        llm_client: LangChain LLM client
        
    Returns:
        TrafficAnalysisGraph instance
    """
    return TrafficAnalysisGraph(input_excel_path, llm_client)
