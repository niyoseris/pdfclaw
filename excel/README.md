# PDFClaw AI Agent for Excel

An Excel spreadsheet with an embedded AI chat agent that launches via VBA macro. Works on both Windows and macOS.

## How It Works

1. Excel spreadsheet contains instructions and a button
2. VBA macro extracts embedded HTML chat app to temp folder
3. HTML opens in default browser with full AI capabilities
4. User enters their Together API key and starts chatting

## Files

| File | Description |
|------|-------------|
| `pdfclaw_excel_agent.xlsx` | Excel spreadsheet with instructions |
| `ai_chat_macro.bas` | VBA module to import into Excel |
| `chat_app.html` | Standalone HTML chat app (for testing) |
| `create_excel_agent.py` | Python script to generate all files |

## Setup Instructions

### Step 1: Enable Macros in Excel

**Windows:**
1. Open Excel
2. Go to **File > Options > Trust Center**
3. Click **Trust Center Settings**
4. Go to **Macro Settings**
5. Select **Enable all macros** or **Disable all macros with notification**
6. Click OK

**macOS:**
1. Open Excel
2. Go to **Excel > Preferences > Security & Privacy**
3. Check **Enable all macros**
4. Restart Excel

### Step 2: Import VBA Module

1. Open `pdfclaw_excel_agent.xlsx`
2. Press **Alt + F11** (Windows) or **Fn + Alt + F11** (Mac) to open VBA Editor
3. Go to **File > Import File**
4. Select `ai_chat_macro.bas`
5. Close VBA Editor

### Step 3: Assign Macro to Button

1. Go to cell B15 (the "LAUNCH AI CHAT" area)
2. Insert a button shape:
   - **Windows:** Developer tab > Insert > Button (Form Control)
   - **macOS:** Insert menu > Shape, then right-click > Assign Macro
3. Draw the button over cell B15
4. Right-click the button > **Assign Macro**
5. Select `LaunchAIChat`
6. Change button text to "🚀 LAUNCH AI CHAT"

### Step 4: Save as Macro-Enabled Workbook

1. **File > Save As**
2. Select **Excel Macro-Enabled Workbook (*.xlsm)**
3. Save as `pdfclaw_excel_agent.xlsm`

## Usage

1. Open `pdfclaw_excel_agent.xlsm`
2. Click "Launch AI Chat" button
3. Browser opens with AI chat interface
4. Enter your [Together API key](https://api.together.xyz/) (free tier available)
5. Start chatting!

## Available Tools

| Tool | Description |
|------|-------------|
| `get_weather` | Get weather for any city |
| `web_search` | Search Wikipedia |
| `fetch_url` | Fetch URL content |
| `calculate` | Math calculations |
| `save_data` | Save data to session |
| `get_data` | Retrieve saved data |
| `list_data` | List all saved data |
| `get_current_time` | Get current time |
| `generate_tweet` | Generate tweet content |
| `generate_content` | Generate articles, emails, etc. |

## Requirements

- Microsoft Excel (Windows or Mac)
- Together API key (free tier available at [api.together.xyz](https://api.together.xyz))

## Security Notes

- **Enable macros only for trusted files**
- **No hardcoded API keys** - Users provide their own
- **API key stored in sessionStorage** - Cleared when browser closes
- **HTML extracted to temp folder** - Automatically cleaned by OS

## Alternative: Standalone HTML

You can also use `chat_app.html` directly:
1. Open `chat_app.html` in any browser
2. Enter your Together API key
3. Start chatting

No Excel or macros required for standalone use.

## Troubleshooting

### Macro doesn't run
- Ensure macros are enabled in Excel settings
- Check if file is saved as `.xlsm` (macro-enabled)

### Button doesn't work
- Make sure macro is assigned to the button
- Try running the macro directly from VBA Editor (F5)

### Browser doesn't open
- Check if temp folder is accessible
- Try running Excel as administrator (Windows)

### VBA Editor doesn't open on Mac
- Use **Fn + Alt + F11** instead of just Alt + F11
- Or go to **Tools > Macro > Visual Basic Editor**

## License

MIT License - part of [PDFClaw](https://github.com/niyoseris/pdfclaw) project.
