import os
import json
from dotenv import load_dotenv
import traceback
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware 

# Import necessary LlamaIndex components
from llama_index.core import Settings
from llama_index.llms.groq import Groq
from llama_index.core.agent import FunctionCallingAgentWorker, AgentRunner
from llama_index.core.tools import FunctionTool
from llama_index.core.agent import AgentChatResponse

# Load environment variables from .env file
load_dotenv()
groq_api_key = os.getenv("GROQ_API_KEY")

# --- FastAPI App Setup ---
app = FastAPI()

# --- CORS Middleware ---
origins = [
    "http://localhost:3000",
    "http://localhost:8080",
    "https://localhost:3000",
    "https://localhost:8080",
    "*", # WARNING: "*" allows *any* origin. Restrict this severely in production!
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Request body model for the API endpoint
class CommandRequest(BaseModel):
    command: str
    workbook_name: str

# --- Instruction Generating Functions ---
# These functions generate the structured instructions for the frontend.
# They do NOT interact with Excel directly.

def generate_read_cell_instruction(workbook_name: str, sheet_name: str, cell_address: str) -> dict:
    """Generates an instruction to read the value from a specified cell in an Excel sheet."""
    return {
        "action": "read_cell",
        "workbook_name": workbook_name,
        "sheet_name": sheet_name,
        "cell_address": cell_address
    }

def generate_write_cell_instruction(workbook_name: str, sheet_name: str, cell_address: str, value: str) -> dict:
    """Generates an instruction to write a string value to a specified cell in an Excel sheet."""
    return {
        "action": "write_cell",
        "workbook_name": workbook_name,
        "sheet_name": sheet_name,
        "cell_address": cell_address,
        "value": value
    }

def generate_read_range_instruction(workbook_name: str, sheet_name: str, range_address: str) -> dict:
    """Generates an instruction to read values from a specified range in an Excel sheet."""
    return {
        "action": "read_range",
        "workbook_name": workbook_name,
        "sheet_name": sheet_name,
        "range_address": range_address
    }

def generate_write_range_instruction(workbook_name: str, sheet_name: str, start_cell_address: str, values_json: str) -> dict:
    """Generates an instruction to write a list of lists (representing rows and columns) to an Excel range, starting from a specified cell."""
    try:
        data = json.loads(values_json)
        if not isinstance(data, list) or not all(isinstance(row, list) for row in data):
             return {
                 "action": "error",
                 "message": "Error: The provided data for writing range is not a valid JSON list of lists."
             }
        return {
            "action": "write_range",
            "workbook_name": workbook_name,
            "sheet_name": sheet_name,
            "start_cell_address": start_cell_address,
            "values": data
        }
    except json.JSONDecodeError:
         return {
             "action": "error",
             "message": "Error: The provided data for writing range is not a valid JSON string."
         }

def generate_get_sheet_names_instruction(workbook_name: str) -> dict:
    """Generates an instruction to get the names of all sheets in the specified workbook."""
    return {
        "action": "get_sheet_names",
        "workbook_name": workbook_name
    }

def generate_clear_range_content_instruction(workbook_name: str, sheet_name: str, range_address: str) -> dict:
    """Generates an instruction to clear the content of a specified range in an Excel sheet."""
    return {
        "action": "clear_range_content",
        "workbook_name": workbook_name,
        "sheet_name": sheet_name,
        "range_address": range_address
    }

def generate_inform_user_instruction(message: str) -> dict:
    """Generates an instruction to simply inform the user with a message."""
    return {
        "action": "inform_user",
        "message": message
    }

# --- NEW Instruction Generating Functions ---

def generate_create_bar_chart_instruction(workbook_name: str, sheet_name: str, data_range: str, chart_title: str, destination_cell: str) -> dict:
    """Generates an instruction to create a bar chart from a specified data range."""
    return {
        "action": "create_bar_chart",
        "workbook_name": workbook_name,
        "sheet_name": sheet_name,
        "data_range": data_range,
        "chart_title": chart_title,
        "destination_cell": destination_cell
    }

def generate_write_formula_instruction(workbook_name: str, sheet_name: str, cell_address: str, formula: str) -> dict:
    """Generates an instruction to write an Excel formula to a specified cell."""
    return {
        "action": "write_formula",
        "workbook_name": workbook_name,
        "sheet_name": sheet_name,
        "cell_address": cell_address,
        "formula": formula
    }

def generate_apply_conditional_formatting_instruction(workbook_name: str, sheet_name: str, range_address: str, condition: str, format_type: str) -> dict:
    """Generates an instruction to apply conditional formatting to a range."""
    # Note: Conditional formatting conditions and format_types will need careful mapping
    # in the frontend Office.js code. The backend just passes the parameters.
    return {
        "action": "apply_conditional_formatting",
        "workbook_name": workbook_name,
        "sheet_name": sheet_name,
        "range_address": range_address,
        "condition": condition, # e.g., "value > 100"
        "format_type": format_type # e.g., "red_fill"
    }

def generate_create_pivot_table_instruction(workbook_name: str, source_sheet: str, source_range: str, dest_sheet: str, dest_cell: str, row_field: str, value_field: str, function: str) -> dict:
    """Generates an instruction to create a pivot table from a source range."""
    # Note: Pivot table creation in Office.js can be complex. This instruction
    # provides the necessary parameters for the frontend to attempt it.
    return {
        "action": "create_pivot_table",
        "workbook_name": workbook_name,
        "source_sheet": source_sheet,
        "source_range": source_range,
        "dest_sheet": dest_sheet,
        "dest_cell": dest_cell,
        "row_field": row_field,
        "value_field": value_field,
        "function": function # e.g., "Sum", "Count"
    }


# --- LlamaIndex Tool Definitions ---
# Wrap the instruction-generating functions as LlamaIndex FunctionTools

read_cell_tool = FunctionTool.from_defaults(
    fn=generate_read_cell_instruction,
    name="generate_read_cell_instruction",
    description="Generates an instruction to read the value from a specified cell in an Excel sheet. Use this tool when the user asks to get the value of a cell (e.g., 'What is in cell A1 on Sheet1?')."
)

write_cell_tool = FunctionTool.from_defaults(
    fn=generate_write_cell_instruction,
    name="generate_write_cell_instruction",
    description="Generates an instruction to write a string value to a specified cell in an Excel sheet. Use this tool when the user asks to write or put data into a cell (e.g., 'Write \"Done\" to cell C1 on Sheet1.')."
)

read_range_tool = FunctionTool.from_defaults(
    fn=generate_read_range_instruction,
    name="generate_read_range_instruction",
    description="Generates an instruction to read values from a specified range in an Excel sheet. Use this tool when the user asks to get data from a range of cells (e.g., 'Read the data in range A1:C5 on Sheet1.')."
)

write_range_tool = FunctionTool.from_defaults(
    fn=generate_write_range_instruction,
    name="generate_write_range_instruction",
    description="Generates an instruction to write a list of lists (representing rows and columns) to an Excel range, starting from a specified cell. The data MUST be provided as a JSON string representing a list of lists (e.g., '[[\"Header1\", \"Header2\"], [\"Data1\", \"Data2\"]]'). Use this tool when the user wants to write multiple values or a table to the sheet (e.g., 'Write the data [[\"Name\", \"Age\"], [\"Alice\", 30]] starting at cell A1 on Sheet2.')."
)

get_sheet_names_tool = FunctionTool.from_defaults(
    fn=generate_get_sheet_names_instruction,
    name="generate_get_sheet_names_instruction",
    description="Generates an instruction to get the names of all sheets in the specified workbook. Use this tool when the user asks for the sheet names or wants to know what sheets are available (e.g., 'What are the names of the sheets?')."
)

clear_range_content_tool = FunctionTool.from_defaults(
    fn=generate_clear_range_content_instruction,
    name="generate_clear_range_content_instruction",
    description="Generates an instruction to clear the content of a specified range in an Excel sheet. Use this tool when the user asks to clear, empty, or delete content from a range or cell (e.g., 'Clear the content of cells A1 to B5 on Sheet1.')."
)

inform_user_tool = FunctionTool.from_defaults(
    fn=generate_inform_user_instruction,
    name="generate_inform_user_instruction",
    description="Generates an instruction to simply inform the user with a message. Use this tool when the user's request doesn't require a specific Excel action but needs a textual response (e.g., asking a general question, or if the request is unclear)."
)

# --- NEW LlamaIndex Tools ---

create_bar_chart_tool = FunctionTool.from_defaults(
    fn=generate_create_bar_chart_instruction,
    name="generate_create_bar_chart_instruction",
    description="Generates an instruction to create a bar chart from a specified data range in an Excel sheet. Use this tool when the user asks to create a bar chart (e.g., 'Create a bar chart for range A1:B5 on Sheet1 titled Sales Data at cell D1'). Requires sheet name, data range, chart title, and a destination cell for the top-left corner of the chart."
)

write_formula_tool = FunctionTool.from_defaults(
    fn=generate_write_formula_instruction,
    name="generate_write_formula_instruction",
    description="Generates an instruction to write an Excel formula to a specified cell. Use this tool when the user asks to write a formula (e.g., 'Write the formula =SUM(A1:A10) into cell A11 on Sheet1'). Requires sheet name, cell address, and the formula string."
)

apply_conditional_formatting_tool = FunctionTool.from_defaults(
    fn=generate_apply_conditional_formatting_instruction,
    name="generate_apply_conditional_formatting_instruction",
    description="Generates an instruction to apply conditional formatting to a range based on a condition and format type. Use this tool when the user asks to format cells based on rules (e.g., 'Highlight values greater than 100 in red in range A1:A10 on Sheet1'). Requires sheet name, range address, a condition description (e.g., 'value > 100'), and a format type description (e.g., 'red_fill')."
)

create_pivot_table_tool = FunctionTool.from_defaults(
    fn=generate_create_pivot_table_instruction,
    name="generate_create_pivot_table_instruction",
    description="Generates an instruction to create a pivot table from a source range on a destination sheet. Use this tool when the user asks to create a pivot table (e.g., 'Create a pivot table from data in A1:C100 on Sheet1 to Sheet2 at A1, grouping by Category and summing Sales'). Requires source sheet, source range, destination sheet, destination cell, row field name, value field name, and aggregation function (e.g., 'Sum', 'Count')."
)


# Define ALL the tools the LlamaIndex agent can use
llamaindex_tools = [
    read_cell_tool,
    write_cell_tool,
    read_range_tool,
    write_range_tool,
    get_sheet_names_tool,
    clear_range_content_tool,
    inform_user_tool,
    # Add the new tools
    create_bar_chart_tool,
    write_formula_tool,
    apply_conditional_formatting_tool,
    create_pivot_table_tool,
]

# --- LlamaIndex Agent Setup ---

# Configure LlamaIndex Settings (including the LLM)
Settings.llm = Groq(model="llama3-70b-8192", api_key=groq_api_key) # Or "llama3-70b-8192"

# Create the Function Calling Agent Worker
agent_worker = FunctionCallingAgentWorker.from_tools(
    llamaindex_tools,
    verbose=True,
    system_prompt="You are an AI assistant that helps users interact with an Excel workbook by generating structured instructions for a frontend application. Based on the user's request and the provided workbook name, determine the appropriate Excel action and its parameters. Use the provided tools to generate a JSON instruction dictionary. Be precise with sheet names and cell/range addresses. For write_range_values, ensure the 'values_json' argument is a valid JSON string representing a list of lists. If the user's request is unclear or does not map to a specific Excel action, use the 'generate_inform_user_instruction' tool. Your final response should be the output of the tool call."
)

# Create the Agent Runner
agent = AgentRunner(agent_worker)


# --- API Endpoint ---
@app.post("/excel-command/")
async def handle_excel_command(request: CommandRequest):
    """
    Receives a natural language command and workbook name from the Excel Add-in,
    processes it using the LlamaIndex agent to generate a structured instruction.
    Returns the structured instruction as a JSON response.
    """
    user_command = request.command
    workbook_name = request.workbook_name

    print(f"Received command for workbook '{workbook_name}': {user_command}")

    try:
        # --- Pass workbook_name to the agent in the query ---
        # Include the workbook name directly in the user query string
        formatted_user_command = f"For the workbook named '{workbook_name}', please {user_command}"
        print(f"Sending formatted command to agent: {formatted_user_command}")

        response_obj: AgentChatResponse = await agent.achat(formatted_user_command) # Use async chat with formatted command

        # --- Extract the actual tool output from sources ---
        instruction = None
        extracted_tool_output = None

        # Iterate through the sources to find the tool output
        if response_obj.sources:
            for source in response_obj.sources:
                print(f"Checking source: {source}")

                tool_output_result = None
                # LlamaIndex versions vary, check common places for tool output
                if hasattr(source, 'raw_output') and source.raw_output is not None:
                    tool_output_result = source.raw_output
                    print(f"Found raw_output in source: {tool_output_result}")

                elif hasattr(source, 'metadata') and source.metadata and 'tool_output' in source.metadata:
                     tool_output_from_metadata = source.metadata['tool_output']
                     if tool_output_from_metadata is not None:
                         tool_output_result = tool_output_from_metadata
                         print(f"Found tool_output in metadata: {tool_output_result}")

                # Check if the extracted result is a dictionary (which our instruction generators return)
                if isinstance(tool_output_result, dict):
                     extracted_tool_output = tool_output_result
                     # Prioritize instructions over the inform_user message if both are called
                     if extracted_tool_output.get("action") != "inform_user":
                          print("Found primary instruction, breaking loop.")
                          break # Found the primary instruction, break the loop
                     else:
                          print("Found inform_user instruction, continuing search for primary instruction.")


        if extracted_tool_output:
            instruction = extracted_tool_output
            print(f"Using extracted tool output as instruction: {instruction}")
        elif response_obj.response:
            # Fallback: if no dictionary tool output was found in sources, use the agent's final text response
            # and wrap it in an 'inform_user' instruction.
            print(f"No dictionary tool output found in sources. Using agent's text response: {response_obj.response}")
            instruction = {"action": "inform_user", "message": response_obj.response}
        else:
            # Final fallback if nothing usable was found
            print("LlamaIndex agent did not produce a structured instruction, tool output, or text response.")
            instruction = {"action": "error", "message": "Backend failed to generate a valid instruction."}

        # --- End Extraction ---

        # Ensure instruction is a dictionary before returning
        if not isinstance(instruction, dict):
             print(f"Final instruction is not a dictionary after extraction: {instruction}")
             instruction = {"action": "error", "message": f"Backend returned unexpected final format after extraction: {str(instruction)}"}


        print(f"Generated instruction: {json.dumps(instruction, indent=2)}")

        # Return the structured instruction dictionary as a JSON response.
        # The frontend will receive this JSON and execute the action.
        return instruction

    except Exception as e:
        print(f"An unexpected error occurred during agent execution: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Backend processing error: {e}")

