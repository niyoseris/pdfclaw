# PDFClaw - AI Agent Embedded in PDF

A fully self-contained AI chat agent embedded inside a PDF file. No external servers, no bridge files - everything is embedded within the PDF itself.

## Features

- **Self-Contained**: Entire AI chat application embedded as HTML file attachment inside PDF
- **Tool Calling**: AI can use tools to fetch weather, search Wikipedia, calculate, store data, generate content
- **Session Storage**: User data persists across the session
- **User API Key**: Users enter their own Together API key (no hardcoded keys)
- **CORS-Friendly APIs**: Uses Wikipedia API, Open-Meteo weather API that work from local files
- **One-Click Launch**: PDF button extracts and opens chat in default browser

## How It Works

1. PDF contains an embedded HTML file (`pdfclaw_chat.html`)
2. "Open AI Chat" button in PDF uses Acrobat JavaScript to extract and launch the HTML
3. HTML opens in browser with full AI chat capabilities
4. User enters their Together API key (stored in sessionStorage)
5. AI responds using tool calling when appropriate

## Tools Available

| Tool | Description | API Used |
|------|-------------|----------|
| `get_weather` | Get current weather for any city | Open-Meteo API |
| `web_search` | Search Wikipedia for information | Wikipedia REST API |
| `fetch_url` | Fetch content from any URL | CORS Proxy |
| `calculate` | Perform mathematical calculations | JavaScript eval |
| `save_data` | Save data to session storage | Browser API |
| `get_data` | Retrieve saved data | Browser API |
| `list_data` | List all saved data | Browser API |
| `get_current_time` | Get current date/time | JavaScript Date |
| `generate_tweet` | Generate tweet-ready content | LLM |
| `generate_content` | Generate articles, emails, code, stories | LLM |

## Installation

```bash
# Clone the repository
git clone https://github.com/niyoseris/pdfclaw.git
cd pdfclaw

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Generate the PDF
python agent_pdf.py
```

## Usage

1. Open `pdfclaw_agent.pdf` in Adobe Acrobat (not Preview, not Chrome)
2. Click the "Open AI Chat" button
3. Your browser will open with the chat interface
4. Enter your [Together API key](https://api.together.xyz/) (free tier available)
5. Start chatting!

### Example Commands

- "What's the weather in Nicosia?"
- "Search Wikipedia for Cyprus"
- "Fetch https://example.com"
- "Calculate 15% of 250"
- "Save my name as John"
- "What's my name?"
- "Generate an article about AI"
- "Write a professional email about meeting"

## Requirements

- Python 3.8+
- reportlab
- pypdf

```bash
pip install reportlab pypdf
```

## Technical Details

### PDF Structure

- Created with ReportLab for visual content
- Modified with pypdf for:
  - Embedded file attachment (HTML chat app)
  - Button annotation with appearance stream
  - JavaScript action for file extraction

### Security Notes

- **No hardcoded API keys** - Users provide their own
- **API key stored in sessionStorage** - Cleared when browser closes
- **CORS limitations** - Some URLs may not be fetchable due to CORS policies
- **Acrobat required** - PDF JavaScript only works in Adobe Acrobat, not Preview or browser viewers

### Why Not Direct HTTP from PDF?

Adobe Acrobat's security sandbox blocks all HTTP requests from document-level JavaScript:
- `Net.HTTP.request` - Blocked
- `SOAP.request` - Blocked  
- `app.trustedFunction` - Blocked

The solution: Embed HTML as file attachment, extract and open in browser where HTTP works normally.

## Project Structure

```
pdfclaw/
├── agent_pdf.py       # Main PDF generation script
├── pdfclaw_agent.pdf  # Generated PDF with embedded chat
├── requirements.txt   # Python dependencies
└── README.md          # This file
```

## API Keys

This project uses [Together API](https://api.together.xyz/) for LLM access:
- Free tier available
- Supports tool calling / function calling
- CORS-friendly (works from local files)

Get your API key at: https://api.together.xyz/

## Limitations

1. **Adobe Acrobat Required**: PDF JavaScript doesn't work in Preview, Chrome, or other PDF viewers
2. **CORS for URL Fetching**: Some websites block CORS requests
3. **Session Storage Only**: Data doesn't persist after browser closes
4. **No Rich Media**: Rich Media annotations deprecated in modern Acrobat

## Contributing

Contributions welcome! Please feel free to submit a Pull Request.

## License

MIT License - feel free to use for any purpose.

## Author

Created by [niyoseris](https://github.com/niyoseris)
