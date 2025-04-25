import os
import json # Needed for writing range data from JSON string
from dotenv import load_dotenv
import xlwings as xw
from langchain_groq import ChatGroq
from langchain.agents import create_tool_calling_agent, AgentExecutor
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain.tools import tool
import traceback # Import traceback for better error logging (optional)


# Load environment variables from .env file
load_dotenv()
# Use environment variable for API Key - Highly Recommended!
groq_api_key = os.getenv("GROQ_API_KEY")

# --- xlwings Functions ---
# We'll use a flexible helper for managing the Excel app instance.

# **Make sure this file exists and the path is correct**
EXCEL_FILE_PATH = "test.xlsx"

def safe_excel_operation(func, **kwargs_from_tool):
    """
    Helper to safely open/close Excel app/book for operations.
    Accepts the core xlwings logic function and keyword arguments intended for it.
    """
    app = None
    book = None
    try:
        # visible=False keeps Excel hidden
        app = xw.App(visible=False, add_book=False)

        # Check if the file exists before opening
        if not os.path.exists(EXCEL_FILE_PATH):
             raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE_PATH}")

        book = app.books.open(EXCEL_FILE_PATH)

        # Call the core logic function, passing the book and keyword arguments
        result = func(book=book, **kwargs_from_tool)

        # Save changes if any were made (safe to call even for read ops)
        book.save()

        return result
    except FileNotFoundError as e:
        print(f"Caught FileNotFoundError: {e}") # Print internal error for debugging
        return f"Error: {e}"
    except Exception as e:
        print(f"Caught Excel Exception: {e}") # Print internal error for debugging
        traceback.print_exc() # Print traceback for detailed debugging (optional)
        # Catch specific xlwings errors if needed, e.g., xw.exceptions.NoSuchObjectError for sheets/ranges
        return f"An Excel error occurred: {e}"
    finally:
        if book:
            try:
                book.close()
            except Exception as close_e:
                print(f"Error closing book: {close_e}")
        if app:
             # It's often better to keep the app running in the background for a while
             # instead of quitting immediately after each operation, especially in a loop.
             # To quit explicitly when done, you could add a cleanup function on script exit.
            pass # app.quit() # Use quit if you want to close the Excel process entirely


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
            # Check if sheet exists
            if sheet_name not in [s.name for s in book.sheets]:
                 return f"Error: Sheet '{sheet_name}' not found."
            sheet = book.sheets[sheet_name]
            value = sheet.range(cell_address).value
            # Convert non-string types to string for returning
            return str(value) if value is not None else ""
       except Exception as e:
            # Catch invalid address errors specifically if possible
            return f"Error reading cell {cell_address} on sheet {sheet_name}: {e}"

    # Call safe_excel_operation, passing the inner function and the tool's arguments
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
             # Check if sheet exists
            if sheet_name not in [s.name for s in book.sheets]:
                 return f"Error: Sheet '{sheet_name}' not found."
            sheet = book.sheets[sheet_name]
            sheet.range(cell_address).value = value
            return f"Successfully wrote '{value}' to cell {cell_address} on sheet {sheet_name}."
        except Exception as e:
             # Catch invalid address errors specifically if possible
            return f"Error writing to cell {cell_address} on sheet {sheet_name}: {e}"

    # Call safe_excel_operation, passing the inner function and the tool's arguments
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
            # Read range values - xlwings returns list of lists for ranges
            values = sheet.range(range_address).value
            # Convert potential None values or non-string types if necessary, then to string
            # Simple conversion to string representation of the list of lists
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
              # Parse the JSON string into a Python list of lists
              data = json.loads(values_json)
              # Basic validation
              if not isinstance(data, list) or not all(isinstance(row, list) for row in data):
                   return "Error: The 'values_json' argument must be a valid JSON string representing a list of lists."

              if sheet_name not in [s.name for s in book.sheets]:
                 return f"Error: Sheet '{sheet_name}' not found."
              sheet = book.sheets[sheet_name]

              # Write the list of lists to the specified range
              sheet.range(start_cell_address).value = data

              return f"Successfully wrote data starting at cell {start_cell_address} on sheet {sheet_name}."
         except json.JSONDecodeError:
              return f"Error: The 'values_json' argument is not a valid JSON string."
         except Exception as e:
              # Handle xlwings specific errors (e.g., invalid address)
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
            # Get names of all sheets
            sheet_names = [sheet.name for sheet in book.sheets]
            return ", ".join(sheet_names) if sheet_names else "No sheets found."
        except Exception as e:
            return f"Error getting sheet names: {e}"

    # Call safe_excel_operation - no specific tool arguments needed beyond the book
    return safe_excel_operation(_get_sheet_names) # No extra kwargs passed


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
            sheet.range(range_address).clear_contents() # Use .clear() to also clear formatting
            return f"Successfully cleared content from range {range_address} on sheet {sheet_name}."
        except Exception as e:
             return f"Error clearing range {range_address} on sheet {sheet_name}: {e}"

    # Call safe_excel_operation, passing the inner function and the tool's arguments
    return safe_excel_operation(_clear_range, sheet_name=sheet_name, range_address=range_address)


# --- LangChain Agent Setup ---

# Define ALL the tools the agent can use
tools = [
    read_cell_value,
    write_cell_value,
    read_range_values,
    write_range_values,
    get_sheet_names,
    clear_range_content,
]

# Initialize the Groq LLM
# !! IMPORTANT: Use environment variable for API key, do not hardcode it here !!
llm = ChatGroq(temperature=0, model_name="llama-3.3-70b-versatile", api_key=groq_api_key) # Use the variable

# Define the prompt for the agent manually
# This prompt guides the LLM to use the available tools
prompt = ChatPromptTemplate.from_messages(
    [
        (
            "system",
            "You are an AI assistant that helps users interact with an Excel file using provided tools. "
            "Be precise with sheet names and cell/range addresses. "
            "When writing ranges, ensure the data is formatted as a JSON string representing a list of lists."
        ),
        # MessagesPlaceholder(variable_name="chat_history"), # Uncomment if adding memory
        ("human", "{input}"),
        MessagesPlaceholder(variable_name="agent_scratchpad"),
    ]
)


# Create the agent
# create_tool_calling_agent is suitable for models that support tool calling
agent = create_tool_calling_agent(llm, tools, prompt)

# Create the agent executor
# handle_parsing_errors=True makes the agent more robust if the LLM outputs malformed tool calls
agent_executor = AgentExecutor(agent=agent, tools=tools, verbose=True, handle_parsing_errors=True)

# --- Main Execution Loop ---

print(f"AI Excel Agent started. Talking to {EXCEL_FILE_PATH}")
print("Make sure Microsoft Excel is installed and the file exists at the specified path.")
print("Type 'exit' to quit.")
print("-" * 20)

while True:
    user_query = input("Enter your Excel request: ")
    if user_query.lower() == 'exit':
        break
    if not user_query.strip():
        continue

    try:
        # Run the agent with the user's query
        response = agent_executor.invoke({"input": user_query})

        print("\nAgent Response:")
        # The response format might vary slightly based on the agent's final step
        # Access 'output' for the final answer or result
        print(response.get('output', 'No output from agent.'))
        print("-" * 20)

    except Exception as e:
        print(f"An unexpected error occurred during agent execution: {e}")
        # Optional: Print traceback here for debugging unexpected errors
        # traceback.print_exc()

print("Agent stopped.")