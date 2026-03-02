"""Executable PDF Agent - Embedded HTML Chat"""

import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.colors import white, black, HexColor
from reportlab.lib.units import cm
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    NameObject, ArrayObject, DictionaryObject,
    TextStringObject, NumberObject,
    DecodedStreamObject
)
import io

# Together API - User enters their own key
MODEL = "ServiceNow-AI/Apriel-1.5-15b-Thinker"

# Self-contained HTML chat app with tool calling
CHAT_HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PDFClaw AI Chat</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #0f0f1a; color: #e0e0e0; height: 100vh; display: flex; flex-direction: column; }}
.header {{ background: #1a1a2e; padding: 12px 16px; border-bottom: 1px solid #2a2a4e; }}
.header h1 {{ font-size: 18px; color: #fff; }}
.header p {{ font-size: 10px; color: #888; margin-top: 2px; }}
.api-key-area {{ background: #16213e; padding: 12px 16px; border-bottom: 1px solid #2a2a4e; display: flex; gap: 8px; align-items: center; }}
.api-key-area input {{ flex: 1; background: #0f0f1a; border: 1px solid #2a2a4e; color: #fff; padding: 8px 12px; border-radius: 4px; font-size: 12px; outline: none; }}
.api-key-area input:focus {{ border-color: #e94560; }}
.api-key-area button {{ background: #e94560; color: #fff; border: none; padding: 8px 16px; border-radius: 4px; font-size: 12px; cursor: pointer; }}
.api-key-area button:hover {{ background: #d63851; }}
.api-key-area .saved {{ color: #7fff7f; font-size: 11px; }}
.chat {{ flex: 1; overflow-y: auto; padding: 12px 16px; display: flex; flex-direction: column; gap: 8px; }}
.msg {{ max-width: 85%; padding: 10px 14px; border-radius: 10px; line-height: 1.4; font-size: 13px; white-space: pre-wrap; word-wrap: break-word; }}
.msg.user {{ background: #e94560; color: #fff; align-self: flex-end; border-bottom-right-radius: 4px; }}
.msg.ai {{ background: #16213e; color: #e0e0e0; align-self: flex-start; border-bottom-left-radius: 4px; }}
.msg.system {{ background: #2a2a4e; color: #888; align-self: center; font-size: 11px; text-align: center; }}
.msg.tool {{ background: #1a3a1a; color: #7fff7f; align-self: flex-start; font-size: 11px; font-family: monospace; }}
.msg.error {{ background: #4a1a1a; color: #ff6b6b; }}
.input-area {{ background: #1a1a2e; padding: 12px 16px; border-top: 1px solid #2a2a4e; display: flex; gap: 8px; }}
.input-area input {{ flex: 1; background: #0f0f1a; border: 1px solid #2a2a4e; color: #fff; padding: 10px 14px; border-radius: 6px; font-size: 13px; outline: none; }}
.input-area input:focus {{ border-color: #e94560; }}
.input-area button {{ background: #e94560; color: #fff; border: none; padding: 10px 20px; border-radius: 6px; font-size: 13px; cursor: pointer; font-weight: 600; }}
.input-area button:hover {{ background: #d63851; }}
.input-area button:disabled {{ background: #555; cursor: not-allowed; }}
.typing {{ color: #888; font-style: italic; }}
</style>
</head>
<body>
<div class="header">
    <h1>PDFClaw AI Agent</h1>
    <p>Model: Apriel-1.5-15b-Thinker | Tool Calling | Session Storage</p>
</div>
<div class="api-key-area">
    <input type="password" id="apiKeyInput" placeholder="Enter your Together API key..." />
    <button onclick="saveApiKey()">Save</button>
    <span class="saved" id="apiKeyStatus"></span>
</div>
<div class="chat" id="chat">
    <div class="msg system">Welcome! Enter your Together API key above. Tools: weather, Wikipedia search, fetch URL, calculate, save/load data, time, tweets, generate content.</div>
</div>
<div class="input-area">
    <input type="text" id="input" placeholder="Type your message..." autofocus />
    <button id="sendBtn" onclick="sendMessage()">Send</button>
</div>
<script>
const API_URL = "https://api.together.xyz/v1/chat/completions";
const MODEL = "{MODEL}";

const chatEl = document.getElementById("chat");
const inputEl = document.getElementById("input");
const sendBtn = document.getElementById("sendBtn");
const apiKeyInput = document.getElementById("apiKeyInput");
const apiKeyStatus = document.getElementById("apiKeyStatus");

// Load saved API key
let API_KEY = sessionStorage.getItem('together_api_key') || '';
if (API_KEY) {{
    apiKeyInput.value = API_KEY;
    apiKeyStatus.textContent = '✓ Saved';
}}

function saveApiKey() {{
    API_KEY = apiKeyInput.value.trim();
    if (API_KEY) {{
        sessionStorage.setItem('together_api_key', API_KEY);
        apiKeyStatus.textContent = '✓ Saved';
    }} else {{
        apiKeyStatus.textContent = '✗ Empty';
    }}
}}

// Session storage for user data
const sessionData = JSON.parse(sessionStorage.getItem('pdfclaw_data') || '{{}}');

// Define tools available to the AI
const tools = [
    {{
        type: "function",
        function: {{
            name: "get_weather",
            description: "Get current weather for a city. Use this for weather queries.",
            parameters: {{
                type: "object",
                properties: {{
                    city: {{ type: "string", description: "City name, e.g. 'Nicosia', 'London'" }}
                }},
                required: ["city"]
            }}
        }}
    }},
    {{
        type: "function",
        function: {{
            name: "web_search",
            description: "Search Wikipedia for information. Returns article summary.",
            parameters: {{
                type: "object",
                properties: {{
                    query: {{ type: "string", description: "Search query for Wikipedia" }}
                }},
                required: ["query"]
            }}
        }}
    }},
    {{
        type: "function",
        function: {{
            name: "fetch_url",
            description: "Fetch content from a specific URL. Returns extracted text.",
            parameters: {{
                type: "object",
                properties: {{
                    url: {{ type: "string", description: "Full URL to fetch, e.g. 'https://example.com'" }}
                }},
                required: ["url"]
            }}
        }}
    }},
    {{
        type: "function",
        function: {{
            name: "calculate",
            description: "Perform mathematical calculations",
            parameters: {{
                type: "object",
                properties: {{
                    expression: {{ type: "string", description: "Math expression to evaluate, e.g. '15*200/100'" }}
                }},
                required: ["expression"]
            }}
        }}
    }},
    {{
        type: "function",
        function: {{
            name: "save_data",
            description: "Save data to session storage for later use",
            parameters: {{
                type: "object",
                properties: {{
                    key: {{ type: "string", description: "Name/key for the data" }},
                    value: {{ type: "string", description: "Value to store" }}
                }},
                required: ["key", "value"]
            }}
        }}
    }},
    {{
        type: "function",
        function: {{
            name: "get_data",
            description: "Retrieve previously saved data from session storage",
            parameters: {{
                type: "object",
                properties: {{
                    key: {{ type: "string", description: "Name/key of the data to retrieve" }}
                }},
                required: ["key"]
            }}
        }}
    }},
    {{
        type: "function",
        function: {{
            name: "list_data",
            description: "List all saved data keys in session storage",
            parameters: {{ type: "object", properties: {{}}, required: [] }}
        }}
    }},
    {{
        type: "function",
        function: {{
            name: "get_current_time",
            description: "Get the current date and time",
            parameters: {{ type: "object", properties: {{}}, required: [] }}
        }}
    }},
    {{
        type: "function",
        function: {{
            name: "generate_tweet",
            description: "Generate a tweet-ready message (max 280 chars)",
            parameters: {{
                type: "object",
                properties: {{
                    topic: {{ type: "string", description: "Topic for the tweet" }},
                    style: {{ type: "string", description: "Style: funny, professional, casual, inspirational" }}
                }},
                required: ["topic"]
            }}
        }}
    }},
    {{
        type: "function",
        function: {{
            name: "generate_content",
            description: "Generate various types of content: articles, emails, code, stories, summaries, etc.",
            parameters: {{
                type: "object",
                properties: {{
                    type: {{ type: "string", description: "Type of content: article, email, code, story, summary, list, poem" }},
                    topic: {{ type: "string", description: "Topic or subject of the content" }},
                    length: {{ type: "string", description: "Length: short, medium, long" }},
                    style: {{ type: "string", description: "Style: formal, casual, professional, creative" }}
                }},
                required: ["type", "topic"]
            }}
        }}
    }}
];

// Tool implementations
async function executeTool(toolName, args) {{
    switch(toolName) {{
        case "get_weather":
            // Open-Meteo API - CORS-friendly, reliable
            try {{
                // First get coordinates for the city
                const geoUrl = `https://geocoding-api.open-meteo.com/v1/search?name=${{encodeURIComponent(args.city)}}&count=1`;
                const geoRes = await fetch(geoUrl);
                if (!geoRes.ok) return `Error finding city: ${{geoRes.status}}`;
                const geoData = await geoRes.json();
                
                if (!geoData.results || geoData.results.length === 0) {{
                    return `City "${{args.city}}" not found`;
                }}
                
                const lat = geoData.results[0].latitude;
                const lon = geoData.results[0].longitude;
                const cityName = geoData.results[0].name;
                
                // Get weather
                const weatherUrl = `https://api.open-meteo.com/v1/forecast?latitude=${{lat}}&longitude=${{lon}}&current_weather=true`;
                const weatherRes = await fetch(weatherUrl);
                if (!weatherRes.ok) return `Error fetching weather: ${{weatherRes.status}}`;
                const weatherData = await weatherRes.json();
                
                const cw = weatherData.current_weather;
                const temp = cw.temperature;
                const wind = cw.windspeed;
                const desc = cw.weathercode <= 3 ? "Clear" : cw.weathercode <= 49 ? "Cloudy/Foggy" : cw.weathercode <= 99 ? "Rain/Drizzle" : "Storm/Snow";
                
                return `Weather in ${{cityName}}: ${{temp}}°C, ${{desc}}, Wind: ${{wind}} km/h`;
            }} catch(e) {{
                return `Error: ${{e.message}}`;
            }}
        
        case "web_search":
            // Wikipedia API - CORS-friendly
            try {{
                const searchUrl = `https://en.wikipedia.org/api/rest_v1/page/summary/${{encodeURIComponent(args.query)}}`;
                const res = await fetch(searchUrl);
                if (!res.ok) {{
                    // Try search API
                    const searchApiUrl = `https://en.wikipedia.org/w/api.php?action=opensearch&search=${{encodeURIComponent(args.query)}}&limit=3&format=json&origin=*`;
                    const searchRes = await fetch(searchApiUrl);
                    if (!searchRes.ok) return `Wikipedia search failed: ${{searchRes.status}}`;
                    const searchData = await searchRes.json();
                    if (searchData[1] && searchData[1].length > 0) {{
                        let results = 'Wikipedia results for "' + args.query + '":\\n';
                        for (let i = 0; i < searchData[1].length; i++) {{
                            results += (i+1) + '. ' + searchData[1][i] + ' - ' + searchData[3][i] + '\\n';
                        }}
                        return results;
                    }}
                    return `No Wikipedia results for "${{args.query}}"`;
                }}
                const data = await res.json();
                return `Wikipedia: ${{data.title}}\n${{data.extract || data.description || 'No summary available'}}\nSource: ${{data.content_urls?.desktop?.page || 'Wikipedia'}}`;
            }} catch(e) {{
                return `Error: ${{e.message}}`;
            }}
        
        case "fetch_url":
            // Fetch URL content via CORS proxy
            try {{
                let url = args.url;
                if (!url.startsWith('http')) url = 'https://' + url;
                
                // Try multiple CORS proxies
                const corsProxies = [
                    `https://api.allorigins.win/raw?url=${{encodeURIComponent(url)}}`,
                    `https://corsproxy.io/?${{encodeURIComponent(url)}}`
                ];
                
                let html = null;
                for (const proxyUrl of corsProxies) {{
                    try {{
                        const res = await fetch(proxyUrl);
                        if (res.ok) {{
                            html = await res.text();
                            break;
                        }}
                    }} catch(e) {{}}
                }}
                
                if (!html) return `Could not fetch ${{url}}. CORS proxies blocked. Try Wikipedia search instead.`;
                
                // Extract text
                const text = html.replace(/<script[^>]*>[\\s\\S]*?<\\/script>/gi, '')
                               .replace(/<style[^>]*>[\\s\\S]*?<\\/style>/gi, '')
                               .replace(/<[^>]+>/g, ' ')
                               .replace(/&nbsp;/g, ' ')
                               .replace(/&amp;/g, '&')
                               .replace(/\\s+/g, ' ')
                               .substring(0, 3000);
                return `Content from ${{url}}:\n${{text}}`;
            }} catch(e) {{
                return `Error: ${{e.message}}`;
            }}
        
        case "calculate":
            try {{
                const result = Function('"use strict"; return (' + args.expression + ')')();
                return `Result: ${{args.expression}} = ${{result}}`;
            }} catch(e) {{
                return `Error calculating: ${{e.message}}`;
            }}
        
        case "save_data":
            sessionData[args.key] = args.value;
            sessionStorage.setItem('pdfclaw_data', JSON.stringify(sessionData));
            return `Saved: ${{args.key}} = "${{args.value}}"`;
        
        case "get_data":
            const value = sessionData[args.key];
            if (value !== undefined) {{
                return `Retrieved: ${{args.key}} = "${{value}}"`;
            }}
            return `No data found for key: ${{args.key}}`;
        
        case "list_data":
            const keys = Object.keys(sessionData);
            if (keys.length === 0) return "No saved data found.";
            return `Saved data keys: ${{keys.join(', ')}}\\nData: ${{JSON.stringify(sessionData, null, 2)}}`;
        
        case "get_current_time":
            const now = new Date();
            return `Current time: ${{now.toLocaleString()}}`;
        
        case "generate_tweet":
            // Return prompt for AI to generate tweet
            return `Generate a tweet about "${{args.topic}}" (style: ${{args.style || 'casual'}}). Keep it under 280 characters. Return ONLY the tweet text.`;
        
        case "generate_content":
            // Return prompt for AI to generate content
            const length = args.length || 'medium';
            const style = args.style || 'professional';
            return `Generate a ${{length}} ${{args.type}} about "${{args.topic}}" in ${{style}} style. Return ONLY the content, no explanations.`;
        
        default:
            return `Unknown tool: ${{toolName}}`;
    }}
}}

let messages = [
    {{ 
        role: "system", 
        content: "You are a helpful AI assistant with access to tools. When you need to search the web, calculate, save/retrieve data, or perform specific tasks, use the appropriate tool. Always be helpful and concise. After using a tool, explain the result to the user naturally."
    }}
];

inputEl.addEventListener("keydown", function(e) {{
    if (e.key === "Enter" && !sendBtn.disabled) sendMessage();
}});

function addMsg(text, cls) {{
    const div = document.createElement("div");
    div.className = "msg " + cls;
    div.textContent = text;
    chatEl.appendChild(div);
    chatEl.scrollTop = chatEl.scrollHeight;
    return div;
}}

async function sendMessage() {{
    const msg = inputEl.value.trim();
    if (!msg) return;
    
    if (!API_KEY) {{
        addMsg("Please enter your Together API key above first!", "system");
        apiKeyInput.focus();
        return;
    }}
    
    addMsg(msg, "user");
    inputEl.value = "";
    sendBtn.disabled = true;
    
    messages.push({{ role: "user", content: msg }});
    
    const typing = addMsg("Thinking...", "ai typing");
    
    try {{
        let response = await callAI();
        typing.textContent = response;
        typing.className = "msg ai";
    }} catch(e) {{
        typing.textContent = "Error: " + e.message;
        typing.className = "msg ai error";
    }}
    
    sendBtn.disabled = false;
    inputEl.focus();
}}

async function callAI() {{
    const res = await fetch(API_URL, {{
        method: "POST",
        headers: {{
            "Content-Type": "application/json",
            "Authorization": "Bearer " + API_KEY
        }},
        body: JSON.stringify({{
            model: MODEL,
            messages: messages,
            tools: tools,
            tool_choice: "auto",
            max_tokens: 1024
        }})
    }});
    
    if (!res.ok) {{
        const errText = await res.text();
        throw new Error(`API Error: ${{res.status}} - ${{errText.substring(0, 100)}}`);
    }}
    
    const data = await res.json();
    const assistantMsg = data.choices[0].message;
    
    // Add assistant message to history
    messages.push(assistantMsg);
    
    // Check if AI wants to use tools
    if (assistantMsg.tool_calls && assistantMsg.tool_calls.length > 0) {{
        // Show tool calls
        for (const toolCall of assistantMsg.tool_calls) {{
            const toolName = toolCall.function.name;
            const args = JSON.parse(toolCall.function.arguments);
            addMsg(`Tool: ${{toolName}}(${{JSON.stringify(args)}})`, "tool");
            
            // Execute tool
            const result = await executeTool(toolName, args);
            addMsg(`Result: ${{result}}`, "tool");
            
            // Add tool result to messages
            messages.push({{
                role: "tool",
                tool_call_id: toolCall.id,
                content: result
            }});
        }}
        
        // Get final response after tool use
        const finalRes = await fetch(API_URL, {{
            method: "POST",
            headers: {{
                "Content-Type": "application/json",
                "Authorization": "Bearer " + API_KEY
            }},
            body: JSON.stringify({{
                model: MODEL,
                messages: messages,
                max_tokens: 1024
            }})
        }});
        
        if (!finalRes.ok) throw new Error("Failed to get final response");
        
        const finalData = await finalRes.json();
        const finalMsg = finalData.choices[0].message;
        messages.push(finalMsg);
        return finalMsg.content;
    }}
    
    return assistantMsg.content || "No response";
}}
</script>
</body>
</html>"""


# JS to extract and open embedded HTML
OPEN_CHAT_JS = 'this.exportDataObject({ cName: "pdfclaw_chat.html", nLaunch: 2 });'


def create_agent_pdf(output_path, title="PDFClaw Agent"):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # Background
    c.setFillColor(HexColor("#1a1a2e"))
    c.rect(0, 0, width, height, fill=True)

    # Header
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 28)
    c.drawString(50, height - 80, "PDFClaw AI Agent")
    c.setFont("Helvetica", 14)
    c.setFillColor(HexColor("#aaaaaa"))
    c.drawString(50, height - 110, "AI Chat fully embedded inside this PDF")

    # Model info
    c.setFillColor(HexColor("#e94560"))
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 160, "Model: Apriel-1.5-15b-Thinker  |  Tool Calling Enabled")

    # Instructions box
    c.setFillColor(HexColor("#16213e"))
    c.roundRect(40, height - 400, width - 80, 220, 10, fill=True)

    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(60, height - 210, "How to use:")

    c.setFont("Helvetica", 12)
    c.setFillColor(HexColor("#cccccc"))
    instructions = [
        "1. Click the 'Open AI Chat' button below",
        "2. Your browser will open with the chat interface",
        "3. Type any message and press Enter or click Send",
        "4. The AI will respond directly",
        "",
        "The chat app is embedded inside this PDF file.",
        "No external files or servers needed.",
    ]
    y = height - 240
    for line in instructions:
        c.drawString(60, y, line)
        y -= 20

    # Features
    c.setFillColor(HexColor("#16213e"))
    c.roundRect(40, height - 560, width - 80, 140, 10, fill=True)

    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(60, height - 440, "Features:")
    c.setFont("Helvetica", 11)
    c.setFillColor(HexColor("#cccccc"))
    features = [
        "- Ask any question, get AI-powered answers",
        "- Draft tweets, emails, messages",
        "- Web research, code generation, translations",
        "- Beautiful dark-mode chat interface",
        "- Conversation history within session",
    ]
    y = height - 465
    for f in features:
        c.drawString(60, y, f)
        y -= 18

    # Footer
    c.setFillColor(HexColor("#666666"))
    c.setFont("Helvetica", 9)
    c.drawString(50, 30, "PDFClaw Agent v2.0 | Everything embedded inside this PDF")

    c.showPage()
    c.save()

    # Build PDF with pypdf
    buffer.seek(0)
    reader = PdfReader(buffer)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    page = writer.pages[0]

    # Embed HTML as file attachment
    html_bytes = CHAT_HTML.encode("utf-8")

    # Create file spec dictionary
    ef_stream = DecodedStreamObject()
    ef_stream.set_data(html_bytes)
    ef_stream[NameObject("/Type")] = NameObject("/EmbeddedFile")
    ef_stream[NameObject("/Subtype")] = NameObject("/text/html")
    ef_stream[NameObject("/Params")] = DictionaryObject({
        NameObject("/Size"): NumberObject(len(html_bytes))
    })
    ef_ref = writer._add_object(ef_stream)

    filespec = DictionaryObject()
    filespec[NameObject("/Type")] = NameObject("/Filespec")
    filespec[NameObject("/F")] = TextStringObject("pdfclaw_chat.html")
    filespec[NameObject("/UF")] = TextStringObject("pdfclaw_chat.html")
    filespec[NameObject("/EF")] = DictionaryObject({
        NameObject("/F"): ef_ref
    })
    filespec_ref = writer._add_object(filespec)

    # Add to Names/EmbeddedFiles
    names_dict = DictionaryObject()
    names_dict[NameObject("/Names")] = ArrayObject([
        TextStringObject("pdfclaw_chat.html"),
        filespec_ref
    ])
    ef_tree = DictionaryObject()
    ef_tree[NameObject("/EmbeddedFiles")] = names_dict
    writer._root_object[NameObject("/Names")] = ef_tree

    # Create "Open AI Chat" button with appearance stream
    btn_dict = DictionaryObject()
    btn_dict[NameObject("/Type")] = NameObject("/Annot")
    btn_dict[NameObject("/Subtype")] = NameObject("/Widget")
    btn_dict[NameObject("/FT")] = NameObject("/Btn")
    btn_dict[NameObject("/Ff")] = NumberObject(65536)
    btn_dict[NameObject("/T")] = TextStringObject("openChatBtn")
    btn_dict[NameObject("/Rect")] = ArrayObject([
        NumberObject(150),
        NumberObject(70),
        NumberObject(int(width - 150)),
        NumberObject(120)
    ])
    
    # Normal appearance stream
    ap_normal = DecodedStreamObject()
    ap_normal.set_data(b"q 1 0 0 1 0 0 cm 0 0 1 rg 0 0 294 50 re f BT /Helv 16 Tf 1 1 1 rg 80 15 Td (Open AI Chat) Tj ET Q")
    ap_normal[NameObject("/Type")] = NameObject("/XObject")
    ap_normal[NameObject("/Subtype")] = NameObject("/Form")
    ap_normal[NameObject("/BBox")] = ArrayObject([NumberObject(0), NumberObject(0), NumberObject(294), NumberObject(50)])
    ap_normal[NameObject("/Resources")] = DictionaryObject({
        NameObject("/Font"): DictionaryObject({
            NameObject("/Helv"): DictionaryObject({
                NameObject("/Type"): NameObject("/Font"),
                NameObject("/Subtype"): NameObject("/Type1"),
                NameObject("/BaseFont"): NameObject("/Helvetica")
            })
        })
    })
    ap_ref = writer._add_object(ap_normal)
    
    btn_dict[NameObject("/AP")] = DictionaryObject({
        NameObject("/N"): ap_ref
    })
    btn_dict[NameObject("/MK")] = DictionaryObject({
        NameObject("/BG"): ArrayObject([NumberObject(0), NumberObject(0), NumberObject(1)]),
        NameObject("/CA"): TextStringObject("Open AI Chat"),
    })
    btn_dict[NameObject("/DA")] = TextStringObject("/Helv 16 Tf 1 1 1 rg")

    # JS action
    js_action = DictionaryObject()
    js_action[NameObject("/S")] = NameObject("/JavaScript")
    js_action[NameObject("/JS")] = TextStringObject(OPEN_CHAT_JS)
    btn_dict[NameObject("/A")] = js_action

    btn_ref = writer._add_object(btn_dict)

    if "/Annots" in page:
        page[NameObject("/Annots")].append(btn_ref)
    else:
        page[NameObject("/Annots")] = ArrayObject([btn_ref])

    # Create AcroForm if needed
    if "/AcroForm" not in writer._root_object:
        acro = DictionaryObject()
        acro[NameObject("/Fields")] = ArrayObject([btn_ref])
        writer._root_object[NameObject("/AcroForm")] = acro
    else:
        acro = writer._root_object["/AcroForm"]
        if "/Fields" in acro:
            acro["/Fields"].append(btn_ref)

    with open(output_path, 'wb') as f:
        writer.write(f)

    return output_path


if __name__ == "__main__":
    create_agent_pdf("pdfclaw_agent.pdf")
    print("PDF created: pdfclaw_agent.pdf")
