# PC-AI-Remote-for-Windows
Remote V42 B - AI-Driven Desktop Automation An Open-Source Alternative to Anthropic's Coworker - Control your Windows PC with natural language.  ? 
# Remote V42 B - AI-Driven Desktop Automation

An Open-Source Alternative to Anthropic's Coworker - Control your Windows PC with natural language.

## ? What is Remote V42 B?

Remote V42 B is a PC remote control program for Windows. It enables computer control through natural language - you describe what you want, and the AI executes it step by step.

**The special feature:** You can connect any LLM via API. Recommended powerful models include:

- Kimi K2.5 (Moonshot AI)
- Claude Sonnet 4.5 (Anthropic) 
- GPT-4o (OpenAI)
- Or any other OpenAI-compatible LLM

## ✨ Features

### 💻 Full System Control via AI

The program can do everything you could via console - just using natural language:

- Read system data (CPU, RAM, disks, network...)
- Read, copy, manage folders
- Delete files **(only with explicit permission and instruction!)**
- Create and run Python programs
- Control the entire PC via AI

Perfect for those afraid of console commands - just tell it in plain language what you want!

### Browser Automation

- Chrome control via Playwright
- Click, type, navigate, scroll
- **Mini-DOM System:** Sends only clickable elements (buttons, links, input fields) to AI - cheaper and more accurate than screenshots!
- **URL-Parameter Trick:** Search directly in URL (`?q=searchterm`) - works reliably on Google, Perplexity, YouTube

### Direct Document Creation

Creates documents directly via Python - no need to open LibreOffice/Word!

- Word documents (.docx)
- Excel spreadsheets (.xlsx)
- PowerPoint presentations (.pptx)

### Desktop Control

- Mouse clicks at any coordinates
- Keyboard input and hotkeys
- Window management via pywinauto
- Screenshot capture with automatic clipboard copy

### Token Tracking

- Real-time display of consumed tokens
- Cost calculation per LLM (customizable price per million tokens)
- **Mini-DOM saves ~70% tokens** compared to screenshots!

### Intelligent Features

- **Context Memory:** Continue conversations or restart with "NEW"
- **Mouse Tracker:** Shows coordinates in real-time (for precise clicks)
- **Error Tracking:** Warns on repeated failures

## Installation

### Prerequisites

```bash
pip install pyautogui playwright requests pillow python-docx openpyxl python-pptx pywinauto pyperclip
playwright install chromium
LLM Setup
    1. Create "llm" folder next to the script
    2. Create text file for your LLM (e.g., kimi-k2.5.txt):
text
URL: https://api.moonshot.ai/v1
API Key: your-api-key-here
LLM Model: kimi-k2.5
Cost per Million: 2.5
Start
bash
python Remote_V42_B.py
Usage Examples
Web Research + Document Creation
text
"Open browser with Perplexity, search for 5 Mini-PCs with technical specs, save as Excel"
→ AI opens Perplexity, reads results, creates .xlsx file
System Information
text
"Show me my system info and save it in a Word document"
→ AI runs PowerShell commands, creates .docx file
Browser Automation
text
"Open YouTube and search for Python tutorials"
→ AI opens youtube.com/results?search_query=Python+Tutorials
Controls
Key/Button	Function
F1	Stop immediately
Send	Execute command
New	Clear context
Screenshot	Capture current screen (→ clipboard)
Last Screenshot AI	Shows what AI last saw
Tracker	Mouse coordinates on/off
LLM Save	Save current LLM settings
Tip: Start a message with NEW to clear context and start fresh.
Available AI Actions
Browser
json
{"action": "browser_start", "url": "https://..."}
{"action": "playwright_click", "selector": "#button-id"}
{"action": "playwright_type", "selector": "input", "text": "Search term"}
{"action": "playwright_get_text", "selector": "body"}
Documents
json
{"action": "create_docx", "path": "C:\\...\\document.docx", "title": "Title", "content": "Text"}
{"action": "create_xlsx", "path": "C:\\...\\table.xlsx", "data": [["A","B"],["1","2"]]}
{"action": "create_pptx", "path": "C:\\...\\presentation.pptx", "slides": [{"title": "...", "content": "..."}]}
Desktop
json
{"action": "mouse_click", "x": 500, "y": 300}
{"action": "key", "key": "Return"}
{"action": "run_commands", "commands": ["start notepad"]}
{"action": "screenshot", "reason": "Check if window open"}
Full System Control via AI
The program can do everything console can - just via natural language:
Read system data
text
"Show me my CPU, RAM, and disk info"
→ AI runs PowerShell/CMD, reads results, presents clearly
Manage files & folders
text
"List all PDF files in my Documents folder"
"Copy all images from Desktop to new 'Backup' folder"
"Delete all temp files" (only with explicit permission!)
Create & run Python programs
text
"Write a Python script that resizes all JPGs in a folder and run it"
→ AI writes code, saves as .py, runs automatically
Complete PC control
    • Start/close programs
    • Change settings
    • Get network info
    • Manage processes
    • And anything else command line can do...
Perfect for those afraid of console - just say what you want in plain language!
⚠️ Security: Program NEVER deletes files without explicit instruction. Reading always allowed, writing/deleting only on command.
? Why Remote V42 B?
Feature	Coworker (Anthropic)	Remote V42 B
Price	Subscription	Just API tokens
Code	Closed Source	Open Source
Runs	Cloud	Local on your machine
LLM	Only Claude	Any LLM
Customizable	No	Fully
Origin
Originally developed for Windows in about 2 weeks across many phases - from first idea through iterations to current V42 B.
    • Linux version: Exists but earlier development stage
    • Mac: Not being created
    • Developed in collaboration between human with 45 years IT experience and AI
License
MIT License - Free to use, modify, and distribute.
Known Limitations
    • Primarily developed for Windows (Linux version available but earlier stage)
    • Mac not supported
    • Chrome must be installed (for browser automation)
    • Complex web pages with dynamic CSS classes can be problematic → use URL-Parameter Trick!

Made by AI & Klaus
