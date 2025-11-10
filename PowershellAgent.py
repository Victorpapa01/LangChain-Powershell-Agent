"""
PowerShell Agent using LangChain v1
This agent can execute PowerShell commands based on natural language instructions.
"""

import os
import subprocess
from dotenv import load_dotenv
from langchain.tools import tool
from langchain.agents import create_agent
from langchain.chat_models import init_chat_model
from langchain_core.rate_limiters import InMemoryRateLimiter

# Rich imports for CLI interface
from rich.console import Console
from rich.panel import Panel
from rich.text import Text
from rich.prompt import Prompt
from rich.markdown import Markdown
from rich.table import Table
from rich import box
from rich.progress import Progress, SpinnerColumn, TextColumn


# Initialize Rich Console
console = Console()

# Load environment variables
load_dotenv()

# Initialize rate limiter for 4 RPM (requests per minute)
# 4 requests per minute = 1 request every 15 seconds
rate_limiter = InMemoryRateLimiter(
    requests_per_second=4/60,  # 4 requests per 60 seconds
    check_every_n_seconds=0.1,  # Check every 100ms
    max_bucket_size=4,  # Maximum burst size
)

# Initialize Google Gemini Flash 2.0 model with rate limiter
model = init_chat_model(
    model="gemini-2.0-flash",
    model_provider="google_genai",
    rate_limiter=rate_limiter,
    max_tokens=500,  # Limit output tokens to 500
    temperature=0.7,
)


@tool
def execute_powershell_command(command: str) -> str:
    """
    Execute a PowerShell command on the terminal and return the output.
    
    Args:
        command: The PowerShell command to execute (e.g., 'Get-ChildItem', 'Get-Process')
    
    Returns:
        The command output or error message
    """
    try:
        # Execute PowerShell command
        result = subprocess.run(
            ["powershell", "-Command", command],
            capture_output=True,
            text=True,
            timeout=30,  # 30 second timeout
            shell=False
        )
        
        if result.returncode == 0:
            output = result.stdout.strip()
            return output if output else "Command executed successfully with no output."
        else:
            error = result.stderr.strip()
            return f"Error executing command: {error}"
            
    except subprocess.TimeoutExpired:
        return "Error: Command execution timed out (30 seconds limit)."
    except Exception as e:
        return f"Error: {str(e)}"


@tool
def search_powershell_command(query: str) -> str:
    """
    Search for PowerShell command usage and help documentation.
    Uses Get-Help to find information about PowerShell commands.
    
    Args:
        query: The command or topic to search for (e.g., 'Get-ChildItem', 'process', 'file operations')
    
    Returns:
        Help documentation and usage examples for the command
    """
    try:
        # Try to get help for the command
        help_command = f"Get-Help {query} -ErrorAction SilentlyContinue"
        result = subprocess.run(
            ["powershell", "-Command", help_command],
            capture_output=True,
            text=True,
            timeout=15,
            shell=False
        )
        
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
        else:
            # If no specific help found, try searching for related commands
            search_command = f"Get-Command *{query}* -ErrorAction SilentlyContinue | Select-Object -First 5 Name, Synopsis"
            result = subprocess.run(
                ["powershell", "-Command", search_command],
                capture_output=True,
                text=True,
                timeout=15,
                shell=False
            )
            
            if result.returncode == 0 and result.stdout.strip():
                return f"Related commands found:\n{result.stdout.strip()}\n\nUse Get-Help <command-name> for more details."
            else:
                return f"No help found for '{query}'. Try using more specific command names like Get-Process, Get-Service, etc."
                
    except subprocess.TimeoutExpired:
        return "Error: Search timed out."
    except Exception as e:
        return f"Error searching for command: {str(e)}"


# Create the agent with tools
tools = [execute_powershell_command, search_powershell_command]

system_prompt = """You are a helpful PowerShell assistant. You can execute PowerShell commands on the user's Windows machine.

IMPORTANT GUIDELINES:
1. Before executing any command, explain what it will do.
2. For destructive operations (delete, remove, stop), ask for confirmation.
3. If you're unsure about a command, use the search_powershell_command tool first.
4. Always provide clear, concise explanations of command outputs.
5. Suggest safer alternatives when appropriate.
6. Keep responses under 500 tokens.

Available tools:
- execute_powershell_command: Execute PowerShell commands on the terminal
- search_powershell_command: Search for PowerShell command help and documentation
"""

# Create agent using LangChain v1
agent = create_agent(
    model=model,
    tools=tools,
    system_prompt=system_prompt
)


def print_banner():
    """Display the ASCII art banner."""
    banner = """
██████╗  ██████╗ ██╗    ██╗███████╗██████╗ ███████╗██╗  ██╗███████╗██╗     ██╗     
██╔══██╗██╔═══██╗██║    ██║██╔════╝██╔══██╗██╔════╝██║  ██║██╔════╝██║     ██║     
██████╔╝██║   ██║██║ █╗ ██║█████╗  ██████╔╝███████╗███████║█████╗  ██║     ██║     
██╔═══╝ ██║   ██║██║███╗██║██╔══╝  ██╔══██╗╚════██║██╔══██║██╔══╝  ██║     ██║     
██║     ╚██████╔╝╚███╔███╔╝███████╗██║  ██║███████║██║  ██║███████╗███████╗███████╗
╚═╝      ╚═════╝  ╚══╝╚══╝ ╚══════╝╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝╚══════╝╚══════╝╚══════╝
                                                                                     
 █████╗  ██████╗ ███████╗███╗   ██╗████████╗                                       
██╔══██╗██╔════╝ ██╔════╝████╗  ██║╚══██╔══╝                                       
███████║██║  ███╗█████╗  ██╔██╗ ██║   ██║                                          
██╔══██║██║   ██║██╔══╝  ██║╚██╗██║   ██║                                          
██║  ██║╚██████╔╝███████╗██║ ╚████║   ██║                                          
╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═╝  ╚═══╝   ╚═╝                                          
    """
    console.print(banner, style="bold bright_magenta")
    console.print("\n")
    
    # Info panel
    info_text = Text()
    info_text.append("Execute PowerShell commands using natural language\n", style="white")
    info_text.append("Powered by ", style="dim")
    info_text.append("Google Gemini Flash 2.0", style="bold green")
    info_text.append(" + ", style="dim")
    info_text.append("LangChain v1", style="bold magenta")
    
    console.print(Panel(
        info_text,
        box=box.ROUNDED,
        border_style="cyan",
        padding=(1, 2)
    ))
    console.print()


def show_config():
    """Display configuration settings."""
    table = Table(show_header=False, box=box.SIMPLE, padding=(0, 2))
    table.add_column(style="cyan bold")
    table.add_column(style="white")
    
    table.add_row("🤖 Model:", "gemini-2.0-flash-exp")
    table.add_row("⚡ Rate Limit:", "4 requests/minute")
    table.add_row("📝 Max Tokens:", "500")
    table.add_row("🔍 LangSmith:", "Enabled" if os.getenv("LANGSMITH_TRACING") == "true" else "Disabled")
    
    console.print(Panel(
        table,
        title="[bold yellow]⚙️  Configuration[/bold yellow]",
        box=box.ROUNDED,
        border_style="yellow",
        padding=(1, 2)
    ))
    console.print()


def show_examples():
    """Display example queries."""
    examples = [
        "List all files in the current directory",
        "Show me the top 5 processes by CPU usage",
        "What is the Get-Process command used for?",
        "Check the Windows version",
        "Show all running services",
    ]
    
    console.print("[bold cyan]💡 Example Queries:[/bold cyan]")
    for i, example in enumerate(examples, 1):
        console.print(f"  [dim]{i}.[/dim] [white]{example}[/white]")
    console.print()


def run_agent(user_input: str):
    """
    Run the PowerShell agent with user input.
    
    Args:
        user_input: Natural language instruction from the user
    """
    # Display user query in a panel
    console.print(Panel(
        f"[bold white]{user_input}[/bold white]",
        title="[bold green]👤 Your Request[/bold green]",
        border_style="green",
        box=box.ROUNDED
    ))
    console.print()
    
    try:
        # Show processing indicator
        with Progress(
            SpinnerColumn("dots"),
            TextColumn("[cyan]Processing your request...[/cyan]"),
            console=console,
            transient=True
        ) as progress:
            progress.add_task("", total=None)
            # Invoke the agent
            result = agent.invoke({"messages": [{"role": "user", "content": user_input}]})
        
        # Extract and print the response
        if "messages" in result and len(result["messages"]) > 0:
            last_message = result["messages"][-1]
            response = last_message.content if hasattr(last_message, 'content') else str(last_message)
            
            # Display response in a panel
            console.print(Panel(
                Markdown(response),
                title="[bold magenta]🤖 Agent Response[/bold magenta]",
                border_style="magenta",
                box=box.ROUNDED,
                padding=(1, 2)
            ))
            console.print()
            return response
        else:
            console.print("[yellow]⚠️  No response generated.[/yellow]\n")
            return None
            
    except Exception as e:
        error_msg = f"Error running agent: {str(e)}"
        console.print(Panel(
            f"[bold red]{error_msg}[/bold red]",
            title="[bold red]❌ Error[/bold red]",
            border_style="red",
            box=box.ROUNDED
        ))
        console.print()
        return None


if __name__ == "__main__":
    # Clear screen (optional)
    os.system('cls' if os.name == 'nt' else 'clear')
    
    # Display banner
    print_banner()
    
    # Show configuration
    show_config()
    
    # Show example queries
    show_examples()
    
    # Instructions
    console.print(Panel(
        "[white]Type your natural language command and press Enter.\n"
        "Commands: [bold cyan]exit[/bold cyan], [bold cyan]quit[/bold cyan], [bold cyan]help[/bold cyan], [bold cyan]config[/bold cyan][/white]",
        title="[bold yellow]📋 Instructions[/bold yellow]",
        border_style="yellow",
        box=box.ROUNDED
    ))
    console.print()
    
    # Interactive loop
    while True:
        try:
            # Get user input with Rich prompt
            user_input = Prompt.ask("[bold green]You[/bold green]").strip()
            
            if not user_input:
                continue
            
            # Handle special commands
            if user_input.lower() in ['exit', 'quit', 'q']:
                console.print("\n[bold cyan]👋 Goodbye! Have a great day![/bold cyan]\n")
                break
            elif user_input.lower() == 'help':
                console.print()
                show_examples()
                continue
            elif user_input.lower() == 'config':
                console.print()
                show_config()
                continue
            
            # Run the agent
            run_agent(user_input)
            
        except KeyboardInterrupt:
            console.print("\n\n[bold cyan]👋 Session interrupted. Goodbye![/bold cyan]\n")
            break
        except Exception as e:
            console.print(f"\n[bold red]❌ Error: {str(e)}[/bold red]\n")
