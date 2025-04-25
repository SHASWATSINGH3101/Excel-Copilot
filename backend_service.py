import os
import json
from dotenv import load_dotenv
import xlwings as xw
from langchain_groq import ChatGroq
from langchain.agents import create_tool_calling_agent, AgentExecutor
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain.tools import tool
import traceback
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel # For request body validation

# Load environment variables from .env file
load_dotenv()
groq_api_key = os.getenv("GROQ_API_KEY")

# --- FastAPI App Setup ---
app = FastAPI()

# Request body model
class CommandRequest(BaseModel):
    command: str

# --- xlwings Functions ---
EXCEL_FILE_PATH = "test.xlsx" # Make sure this file exists and path is correct

def safe_excel_operation(func, **kwargs_from_tool):
    """
    Helper to safely open/close Excel app/book for operations.
    Accepts the core xlwings logic function and keyword arguments intended for it.
    """
    app = None
    book = None
    try:
        app = xw.App(visible=False, add_book=False)

        if not os.path.exists(EXCEL_FILE_PATH):
             raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE_PATH}")

        book = app.books.open(EXCEL_FILE_PATH)

        result = func(book=book, **kwargs_from_tool)

        book.save()

        return {"status": "success", "result": result} # Return structured result

    except FileNotFoundError as e:
        print(f"Caught FileNotFoundError: {e}")
        return {"status": "error", "message": f"Error: {e}"}
    except Exception as e:
        print(f"Caught Excel Exception: {e}")
        traceback.print_exc()
        return {"status": "error", "message": f"An Excel error occurred: {e}"}
    finally:
        if book:
            try:
                book.close()
            except Exception as close_e:
                print(f"Error closing book: {close_e}")
        if app:
            pass


# --- Tool Definitions ---

@tool
def read_cell_value(sheet_name: str, cell_address: str) -> str:
    """Reads the value from a specified cell in an Excel sheet.
    Use this tool when the user asks to get the value of a cell (e.g., 'What is in cell A1 on Sheet1?').
    Args:
        sheet_name: The name of the sheet (e.g., 'Sheet1'). Case-sensitive.
        cell_address: The address of the cell (e.g., 'A1', 'B5').
    Returns:
        The value of the cell as a string, or an error message.
    """
    def _read_cell(book, sheet_name: str, cell_address: str):
       try:
            if sheet_name not in [s.name for s in book.sheets]:
                 return f"Error: Sheet '{sheet_name}' not found."
            sheet = book.sheets[sheet_name]
            value = sheet.range(cell_address).value
            return str(value) if value is not None else ""
       except Exception as e:
            return f"Error reading cell {cell_address} on sheet {sheet_name}: {e}"

    return safe_excel_operation(_read_cell, sheet_name=sheet_name, cell_address=cell_address)


@tool
def write_cell_value(sheet_name: str, cell_address: str, value: str) -> str:
    """Writes a string value to a specified cell in an Excel sheet.
    Use this tool when the user asks to write or put data into a cell (e.g., 'Write "Done" to cell C1 on Sheet1.').
    Args:
        sheet_name: The name of the sheet (e.g., 'Sheet1'). Case-sensitive.
        cell_address: The address of the cell (e.g., 'B2', 'C10').
        value: The string value to write to the cell. The value should be the exact text provided by the user.
    Returns:
        A success message or an error message.
    """
    def _write_cell(book, sheet_name: str, cell_address: str, value: str):
        try:
            if sheet_name not in [s.name for s in book.sheets]:
                 return f"Error: Sheet '{sheet_name}' not found."
            sheet = book.sheets[sheet_name]
            sheet.range(cell_address).value = value
            return f"Successfully wrote '{value}' to cell {cell_address} on sheet {sheet_name}."
        except Exception as e:
            return f"Error writing to cell {cell_address} on sheet {sheet_name}: {e}"

    return safe_excel_operation(_write_cell, sheet_name=sheet_name, cell_address=cell_address, value=value)


@tool
def read_range_values(sheet_name: str, range_address: str) -> str:
    """Reads values from a specified range in an Excel sheet and returns them as a string representation of a list of lists.
    Use this tool when the user asks to get data from a range of cells (e.g., 'Read the data in range A1:C5 on Sheet1.').
    Args:
        sheet_name: The name of the sheet (e.g., 'Sheet1'). Case-sensitive.
        range_address: The address of the range (e.g., 'A1:C5', 'B:D').
    Returns:
        A string representation of a list of lists containing the range values, or an error message.
    """
    def _read_range(book, sheet_name: str, range_address: str):
        try:
            if sheet_name not in [s.name for s in book.sheets]:
                 return f"Error: Sheet '{sheet_name}' not found."
            sheet = book.sheets[sheet_name]
            values = sheet.range(range_address).value
            return str(values)
        except Exception as e:
            return f"Error reading range {range_address} on sheet {sheet_name}: {e}"

    return safe_excel_operation(_read_range, sheet_name=sheet_name, range_address=range_address)


@tool
def write_range_values(sheet_name: str, start_cell_address: str, values_json: str) -> str:
    """Writes a list of lists (representing rows and columns) to an Excel range, starting from a specified cell.
    The data MUST be provided as a JSON string representing a list of lists (e.g., '[["Header1", "Header2"], ["Data1", "Data2"]]').
    Use this tool when the user wants to write multiple values or a table to the sheet (e.g., 'Write the data [["Name", "Age"], ["Alice", 30]] starting at cell A1 on Sheet2.').
    Args:
        sheet_name: The name of the sheet (e.g., 'Sheet1'). Case-sensitive.
        start_cell_address: The address of the top-left cell of the range (e.g., 'A1', 'C5').
        values_json: A JSON string representing the list of lists to write. Example: '[["Header1", "Header2"], ["Data1", "Data2"]]'
    Returns:
        A success message or an error message.
    """
    def _write_range(book, sheet_name: str, start_cell_address: str, values_json: str):
         try:
              data = json.loads(values_json)
              if not isinstance(data, list) or not all(isinstance(row, list) for row in data):
                   return "Error: The 'values_json' argument must be a valid JSON string representing a list of lists."

              if sheet_name not in [s.name for s in book.sheets]:
                 return f"Error: Sheet '{sheet_name}' not found."
              sheet = book.sheets[sheet_name]
              sheet.range(start_cell_address).value = data
              return f"Successfully wrote data starting at cell {start_cell_address} on sheet {sheet_name}."
         except json.JSONDecodeError:
              return f"Error: The 'values_json' argument is not a valid JSON string."
         except Exception as e:
              return f"An Excel error occurred during write range to {start_cell_address} on {sheet_name}: {e}"

    return safe_excel_operation(_write_range, sheet_name=sheet_name, start_cell_address=start_cell_address, values_json=values_json)


@tool
def get_sheet_names() -> str:
    """Gets the names of all sheets in the Excel file.
    Use this tool when the user asks for the sheet names or wants to know what sheets are available (e.g., 'What are the names of the sheets?').
    Returns:
        A comma-separated string of sheet names, or an error message.
    """
    def _get_sheet_names(book):
        try:
            sheet_names = [sheet.name for sheet in book.sheets]
            return ", ".join(sheet_names) if sheet_names else "No sheets found."
        except Exception as e:
            return f"Error getting sheet names: {e}"

    return safe_excel_operation(_get_sheet_names)


@tool
def clear_range_content(sheet_name: str, range_address: str) -> str:
    """Clears the content of a specified range in an Excel sheet.
    Use this tool when the user asks to clear, empty, or delete content from a range or cell (e.g., 'Clear the content of cells A1 to B5 on Sheet1.').
    Args:
        sheet_name: The name of the sheet (e.g., 'Sheet1'). Case-sensitive.
        range_address: The address of the range (e.g., 'A1:B5', 'C:C', '1:1', 'D5').
    Returns:
        A success message or an error message.
    """
    def _clear_range(book, sheet_name: str, range_address: str):
        try:
            if sheet_name not in [s.name for s in book.sheets]:
                 return f"Error: Sheet '{sheet_name}' not found."
            sheet = book.sheets[sheet_name]
            sheet.range(range_address).clear_contents()
            return f"Successfully cleared content from range {range_address} on sheet {sheet_name}."
        except Exception as e:
             return f"Error clearing range {range_address} on sheet {sheet_name}: {e}"

    return safe_excel_operation(_clear_range, sheet_name=sheet_name, range_address=range_address)


# --- LangChain Agent Setup ---
tools = [
    read_cell_value,
    write_cell_value,
    read_range_values,
    write_range_values,
    get_sheet_names,
    clear_range_content,
]

llm = ChatGroq(temperature=0, model_name="llama3-8b-8192", api_key=groq_api_key)

prompt = ChatPromptTemplate.from_messages(
    [
        (
            "system",
            "You are an AI assistant that helps users interact with an Excel file using provided tools. "
            "Be precise with sheet names and cell/range addresses. "
            "When writing ranges, ensure the data is formatted as a JSON string representing a list of lists. "
            "After using a tool and getting a result, formulate a clear and concise answer for the user based on the tool's output. "
            "If the user's request is fulfilled by the tool output, provide that output as your final answer."
        ),
        ("human", "{input}"),
        MessagesPlaceholder(variable_name="agent_scratchpad"),
    ]
)

agent = create_tool_calling_agent(llm, tools, prompt)

# Removed verbose=True for cleaner API output
agent_executor = AgentExecutor(agent=agent, tools=tools, handle_parsing_errors=True)


# --- API Endpoint ---
@app.post("/excel-command/")
async def handle_excel_command(request: CommandRequest):
    """
    Receives a natural language command and processes it using the LangChain agent.
    Returns the agent's final response or tool output.
    """
    user_command = request.command

    print(f"Received command: {user_command}") # Log received command

    try:
        # Invoke the agent with the user command
        # The agent will decide which tool to use and execute it
        response = agent_executor.invoke({"input": user_command})

        # The agent's final output will be in the 'output' key
        agent_output = response.get('output', 'The agent did not produce a final output.')

        print(f"Agent output: {agent_output}") # Log agent output

        # Return the agent's output as a JSON response
        # In a real add-in, you might need to structure this output
        # more carefully into "JSON Instructions" as per your flowchart.
        # For now, we return the final text output.
        return {"response": agent_output}

    except Exception as e:
        print(f"An error occurred during agent execution: {e}") # Log error
        traceback.print_exc() # Print traceback for debugging
        raise HTTPException(status_code=500, detail=f"Backend processing error: {e}")

# --- How to Run ---
# 1. Save the code as a Python file (e.g., main.py).
# 2. Make sure you have uvicorn installed (`pip install uvicorn[standard] fastapi`).
# 3. Run the service from your terminal: `uvicorn main:app --reload`
# 4. The service will start, usually at http://127.0.0.1:8000.
# 5. Your Excel Add-in frontend would then send POST requests to http://127.0.0.1:8000/excel-command/