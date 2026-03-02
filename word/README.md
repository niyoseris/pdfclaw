# PDFClaw AI Agent for Word

A Word document with an embedded AI chat agent that launches via VBA macro.

## How It Works

1. Word document contains instructions and a button placeholder
2. VBA macro extracts embedded HTML chat app to temp folder
3. HTML opens in default browser with full AI capabilities
4. User enters their Together API key and starts chatting

## Files

| File | Description |
|------|-------------|
| `pdfclaw_word_agent.docx` | Word document with instructions |
| `chat_app.html` | Standalone HTML chat app (for testing) |
| `ai_chat_macro.bas` | VBA module to import into Word |
| `create_word_agent.py` | Python script to generate all files |

## Setup Instructions

### Step 1: Enable Macros in Word

1. Open Word
2. Go to **File > Options > Trust Center**
3. Click **Trust Center Settings**
4. Go to **Macro Settings**
5. Select **Enable all macros** (or **Disable all macros with notification**)
6. Click OK

### Step 2: Import VBA Module

1. Open `pdfclaw_word_agent.docx`
2. Press **Alt + F11** to open VBA Editor
3. Go to **File > Import File**
4. Select `ai_chat_macro.bas`
5. Close VBA Editor

### Step 3: Add Button

1. Go to **Developer tab** (enable it first: File > Options > Customize Ribbon > check Developer)
2. Click **Button (Form Control)** or use **Insert > Shapes** for a button shape
3. Draw the button on the document
4. Right-click the button > **Assign Macro**
5. Select `LaunchAIChat`
6. Change button text to "Launch AI Chat"

### Step 4: Save as Macro-Enabled Document

1. **File > Save As**
2. Select **Word Macro-Enabled Document (*.docm)**
3. Save as `pdfclaw_word_agent.docm`

## Usage

1. Open `pdfclaw_word_agent.docm`
2. Click "Launch AI Chat" button
3. Browser opens with AI chat interface
4. Enter your [Together API key](https://api.together.xyz/)
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

- Microsoft Word (Windows or Mac with VBA support)
- Together API key (free tier available)

## Security Notes

- **Enable macros only for trusted documents**
- **No hardcoded API keys** - Users provide their own
- **API key stored in sessionStorage** - Cleared when browser closes
- **HTML extracted to temp folder** - Deleted on system restart

## Alternative: Standalone HTML

You can also use `chat_app.html` directly:
1. Open `chat_app.html` in any browser
2. Enter your Together API key
3. Start chatting

No Word or macros required for standalone use.

## Troubleshooting

### Macro doesn't run
- Ensure macros are enabled
- Check if document is saved as `.docm` (macro-enabled)

### Button doesn't appear
- Enable Developer tab in Word options
- Use Insert > Shapes to create a button manually

### Browser doesn't open
- Check if temp folder is accessible
- Try running Word as administrator

## License

MIT License - part of [PDFClaw](https://github.com/niyoseris/pdfclaw) project.
