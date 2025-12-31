"""
Main Entry Point for GA4 LangGraph Traffic Analysis
Executes the complete agentic AI workflow.

SETUP INSTRUCTIONS:
===================

1. CREATE VIRTUAL ENVIRONMENT:
   
   macOS / Linux:
   --------------
   python3 -m venv venv
   source venv/bin/activate
   
   Windows:
   --------
   python -m venv venv
   venv\\Scripts\\activate

2. INSTALL DEPENDENCIES:
   pip install -r requirements.txt

3. CONFIGURE .ENV FILE:
   - Set your LLM provider (openai, openrouter, or ollama)
   - Add your API key
   - Verify file paths

4. RUN THE PROJECT:
   python main.py
"""

import os
import sys
from dotenv import load_dotenv
from openai import OpenAI
from langchain_openai import ChatOpenAI

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from graph.traffic_analysis_graph import create_traffic_analysis_graph


def initialize_llm_client():
    """
    Initialize LLM client based on .env configuration.
    
    Returns:
        LangChain LLM client instance
    """
    provider = os.getenv('LLM_PROVIDER', 'openai').lower()
    
    print(f"Initializing LLM Provider: {provider.upper()}")
    
    if provider == 'openai':
        api_key = os.getenv('OPENAI_API_KEY')
        if not api_key or api_key == 'your_key_here':
            raise ValueError("OPENAI_API_KEY not set in .env file")
        
        llm = ChatOpenAI(
            model="gpt-4",
            temperature=0.3,
            api_key=api_key
        )
        print("[+] OpenAI GPT-4 initialized")
        
    elif provider == 'openrouter':
        api_key = os.getenv('OPENROUTER_API_KEY')
        if not api_key or api_key == 'your_key_here':
            raise ValueError("OPENROUTER_API_KEY not set in .env file")
        
        llm = ChatOpenAI(
            model="mistralai/devstral-2512:free",
            temperature=0.3,
            api_key=api_key,
            base_url="https://openrouter.ai/api/v1",
            max_tokens=500
        )
        print("[+] OpenRouter initialized with Mistral Devstral Free model")
        
    elif provider == 'ollama':
        from langchain_community.chat_models import ChatOllama
        
        model_name = os.getenv('OLLAMA_MODEL', 'llama3')
        llm = ChatOllama(
            model=model_name,
            temperature=0.3
        )
        print(f"[+] Ollama ({model_name}) initialized")
        
    else:
        raise ValueError(f"Unsupported LLM_PROVIDER: {provider}")
    
    return llm


def validate_environment():
    """
    Validate that all required environment variables and files are present.
    
    Raises:
        ValueError: If configuration is invalid
        FileNotFoundError: If required files are missing
    """
    print("\n" + "="*60)
    print("VALIDATING ENVIRONMENT")
    print("="*60)
    
    # Check .env file loaded
    if not os.getenv('LLM_PROVIDER'):
        raise ValueError(".env file not loaded or LLM_PROVIDER not set")
    
    print("[+] Environment variables loaded")
    
    # Check input file
    input_path = os.getenv('INPUT_EXCEL_PATH')
    if not input_path:
        raise ValueError("INPUT_EXCEL_PATH not set in .env")
    
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input Excel file not found: {input_path}")
    
    print(f"[+] Input file exists: {input_path}")
    
    # Check output path directory
    output_path = os.getenv('OUTPUT_EXCEL_PATH')
    if not output_path:
        raise ValueError("OUTPUT_EXCEL_PATH not set in .env")
    
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"[+] Created output directory: {output_dir}")
    else:
        print(f"[+] Output directory exists: {output_dir}")
    
    print("="*60 + "\n")


def main():
    """
    Main execution function.
    """
    print("\n" + "="*60)
    print("GA4 LANGGRAPH TRAFFIC ANALYSIS - AGENTIC AI SYSTEM")
    print("="*60 + "\n")
    
    # Load environment variables
    load_dotenv()
    print("[+] Loaded .env configuration")
    
    # Validate environment
    try:
        validate_environment()
    except (ValueError, FileNotFoundError) as e:
        print(f"[ERROR] Environment validation failed: {e}")
        print("\nPlease check your .env file and input data.")
        sys.exit(1)
    
    # Initialize LLM client
    try:
        llm_client = initialize_llm_client()
    except Exception as e:
        print(f"[ERROR] Failed to initialize LLM: {e}")
        print("\nPlease check your API keys and LLM configuration.")
        sys.exit(1)
    
    # Get file paths
    input_excel_path = os.getenv('INPUT_EXCEL_PATH')
    output_excel_path = os.getenv('OUTPUT_EXCEL_PATH')
    
    print(f"\nInput File: {input_excel_path}")
    print(f"Output File: {output_excel_path}\n")
    
    # Create and run LangGraph workflow
    try:
        print("="*60)
        print("BUILDING LANGGRAPH WORKFLOW")
        print("="*60 + "\n")
        
        # Create graph
        graph = create_traffic_analysis_graph(input_excel_path, llm_client)
        
        print("[+] LangGraph workflow built successfully")
        print("[+] Agents initialized:")
        print("  1. ExcelReaderAgent")
        print("  2. SectionDetectorAgent")
        print("  3. MetricsCalculatorAgent")
        print("  4. SummaryGeneratorAgent")
        print("  5. ExcelWriterAgent")
        
        # Execute workflow
        final_state = graph.run(recursion_limit=100)
        
        # Save output workbook
        print("\n" + "="*60)
        print("SAVING OUTPUT")
        print("="*60 + "\n")
        
        workbook = final_state.get('workbook')
        if workbook:
            workbook.save(output_excel_path)
            print(f"[+] Output saved to: {output_excel_path}")
        else:
            print("[!] Warning: No workbook in final state")
        
        # Print summary
        print("\n" + "="*60)
        print("EXECUTION SUMMARY")
        print("="*60)
        
        sections_processed = len(final_state.get('sections', []))
        print(f"[+] Sections processed: {sections_processed}")
        print(f"[+] YOY percentages calculated")
        print(f"[+] LM percentages calculated")
        print(f"[+] Executive summaries generated")
        print(f"[+] Results written to Excel")
        
        print("\n" + "="*60)
        print("WORKFLOW COMPLETED SUCCESSFULLY")
        print("="*60 + "\n")
        
        print(f"Open the output file to view results:")
        print(f"   {output_excel_path}\n")
        
    except Exception as e:
        print(f"\n[ERROR] ERROR DURING EXECUTION:")
        print(f"   {str(e)}")
        print("\nStack trace:")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
