# PowerShell Agent with LangChain v1

A powerful LangChain v1 agent that executes PowerShell commands on Windows based on natural language instructions.

## Features

- **Beautiful CLI Interface**: Polished terminal UI with ASCII art banner, colors, and formatted panels
- **Natural Language Interface**: Give commands in plain English, and the agent translates them to PowerShell
- **Two Custom Tools**:
  1. `execute_powershell_command`: Executes PowerShell commands on your terminal
  2. `search_powershell_command`: Searches PowerShell help documentation and command usage
- **Google Gemini Flash 2.0**: Uses the latest Gemini model for intelligent command generation
- **Rate Limiting**: Configured to stay under 4 requests per minute
- **Token Limit**: Max output tokens limited to 500 for concise responses
- **Rich Terminal Features**:
  - ASCII art banner
  - Colored output with syntax highlighting
  - Bordered panels for organized information
  - Progress spinners during processing
  - Markdown rendering for responses
  - Interactive prompts
- **Built-in Commands**: 
  - `help` - Show example queries
  - `config` - Display current configuration
  - `exit/quit` - Exit the application
- **Safety Features**: 
  - Explains commands before execution
  - Warns about destructive operations
  - 30-second timeout for command execution
- **LangSmith Tracing**: Integrated tracing for debugging (configured in .env)

## Prerequisites

- Python 3.8+
- Windows operating system with PowerShell
- Google API key for Gemini

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Ensure your `.env` file contains the Google API key:
```env
GOOGLE_API_KEY="your-api-key-here"
```

## Usage

### Interactive Mode

Run the agent in interactive mode:
```bash
python PowershellAgent.py
```

Then type your natural language commands or built-in commands:
```
You: List all files in the current directory
You: Show me the running processes
You: What is the Get-Process command used for?
You: help          # Show example queries
You: config        # Display configuration
You: exit          # Exit the application
```

### Programmatic Usage

```python
from PowershellAgent import run_agent

# Execute a single command
run_agent("List all running Python processes")
```

## CLI Interface

The agent features a beautiful terminal interface powered by the Rich library:

### Visual Elements
- **ASCII Art Banner**: Eye-catching "POWERSHELL AGENT" banner in cyan
- **Information Panel**: Shows the agent description and tech stack
- **Configuration Table**: Displays model, rate limits, and settings
- **Example Queries**: Quick-start examples for new users
- **Colored Panels**: 
  - Green panels for user requests
  - Magenta panels for agent responses
  - Yellow panels for instructions
  - Red panels for errors
- **Progress Spinners**: Visual feedback during processing
- **Markdown Rendering**: Formatted code blocks and text in responses

### Built-in Commands
- `help` - Display example queries
- `config` - Show current configuration settings
- `exit` or `quit` - Exit the application gracefully

## Example Queries

- "List all files in the current directory"
- "Show me the top 5 processes by CPU usage"
- "Get information about the Get-Service command"
- "Check the Windows version"
- "Show all running services"
- "What command do I use to check disk space?"

## Configuration

### Rate Limiting
The agent is configured for 4 requests per minute (RPM):
```python
rate_limiter = InMemoryRateLimiter(
    requests_per_second=4/60,  # 4 requests per 60 seconds
    check_every_n_seconds=0.1,
    max_bucket_size=4,
)
```

### Max Tokens
Output is limited to 500 tokens for concise responses:
```python
model = init_chat_model(
    model="gemini-2.0-flash-exp",
    max_tokens=500,
    ...
)
```

## Safety Notes

⚠️ **Important**: This agent can execute PowerShell commands on your system. 

- The agent will explain commands before execution
- Be cautious with destructive operations (delete, remove, stop)
- Commands have a 30-second timeout
- Review the agent's explanation before allowing execution

## Architecture

```
PowershellAgent.py
├── Environment Setup (.env loading)
├── Rate Limiter Configuration (4 RPM)
├── Model Initialization (Gemini Flash 2.0, 500 max tokens)
├── Custom Tools
│   ├── execute_powershell_command (executes commands)
│   └── search_powershell_command (searches help docs)
├── Agent Creation (create_agent with system prompt)
└── Interactive Loop (user input → agent → tool execution → response)
```

## Troubleshooting

### Rate Limit Issues
If you hit rate limits, the agent will automatically wait. The rate limiter ensures you stay under 4 RPM.

### Command Timeouts
Commands are limited to 30 seconds. For long-running operations, consider breaking them into smaller steps.

### API Key Issues
Ensure your `.env` file contains a valid `GOOGLE_API_KEY`:
```bash
GOOGLE_API_KEY="AIzaSy..."
```

### Import Errors
If you encounter import errors, ensure all dependencies are installed:
```bash
pip install -r requirements.txt --upgrade
```

## LangSmith Tracing

The agent automatically uses LangSmith tracing (configured in your `.env`):
```env
LANGSMITH_TRACING=true
LANGSMITH_ENDPOINT=https://api.smith.langchain.com
LANGSMITH_API_KEY=your-langsmith-key
LANGSMITH_PROJECT=langchain-tutorials
```

View traces at [smith.langchain.com](https://smith.langchain.com)

## License

This project is for educational purposes.
