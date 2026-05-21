# Ollama Outlook AI

Local AI email assistant for Outlook 2013 — powered by Ollama.

## Description

Ollama Outlook AI is a VBA macro for Microsoft Outlook 2013+ that adds AI-powered email assistance using locally-running Ollama models. It is a recoded version of Intel's [key-assist-outlook](https://github.com/paulchi-intel/key-assist-outlook) that replaces the proprietary Intel ExpertGPT API with Ollama's OpenAI-compatible local endpoint.

No cloud dependency. No API keys. No data leaves your machine.

## Features

- Fully local and offline — works without internet
- No API key required — Ollama runs on your own computer
- Compatible with any Ollama model (llama3.2, mistral, gemma, phi, etc.)
- Privacy-preserving — email content never leaves your machine
- Works with Outlook 2013 and later
- Configurable system prompt for custom AI behavior
- Model selection dropdown populated from your installed Ollama models
- Adjustable timeout for slower models
- Toolbar button in email compose windows for one-click AI assistance

## Prerequisites

1. **Ollama** installed and running — download from [ollama.com](https://ollama.com)
2. **At least one model pulled** — run `ollama pull llama3.2:3b` (2GB, runs well on CPU)
3. **Microsoft Outlook 2013 or later** (any edition)
4. **Windows** with VBA support (built into Outlook)

## Installation

1. Open Outlook and press **Alt+F11** to open the VBA editor
2. In the menu bar, go to **File > Import File**
3. Import **both** files (order doesn't matter):
   - `ollama-outlook-ai.bas` (main module)
   - `ollama-events.cls` (event handler class)
4. Close the VBA editor and restart Outlook
5. Go to **Tools > Macros > Ollama_Initialize** and click **Run** (only needed once)

The **Ollama AI** toolbar will now appear at the top of email compose windows.

## Configuration

To configure the addon:

1. In an email compose window, click the **Ollama AI** button dropdown
2. Click **Settings**
3. A configuration form appears where you can:
   - **Select Model**: Choose from your installed Ollama models (fetched automatically)
   - **Custom Prompt**: Edit the system prompt to tailor AI responses
   - **Timeout**: Adjust API timeout (default 120s, increase for slower models)

All settings are saved in Windows Registry under `HKCU\Software\OllamaAI`.

## Usage

1. Open a new email or reply to an existing one
2. Write or review the email content
3. Click the **Ollama AI** button in the toolbar
4. The AI analyzes your email and inserts a response/suggestion at the cursor position

The button works on:
- New email composition
- Reply and Reply All
- Forwarded messages

## How It Works

1. When you click "Ollama AI", the macro reads the current email's subject and body
2. It sends the content to Ollama's OpenAI-compatible API at `http://localhost:11434/v1/chat/completions`
3. Ollama processes the request using your selected model
4. The AI-generated response is inserted into your email at the cursor position

All communication is via HTTP to `localhost` — no data travels over the network.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "Ollama is not running" | Open a terminal and run `ollama list` to verify. Start Ollama from Start Menu if not running. |
| No models in dropdown | Run `ollama pull llama3.2:3b` in terminal, then reopen Settings. |
| Request times out | Increase timeout in Settings (120s+). Smaller models respond faster. |
| Button not appearing | Run Tools > Macros > Ollama_Initialize once, then restart Outlook. |
| VBA security warning | Go to File > Options > Trust Center > Trust Center Settings > Macro Settings > Enable all macros. |

## Technical Details

- **File to import**: `ollama-outlook-ai.bas` (VBA standard module)

- **Language**: VBA (Visual Basic for Applications)
- **HTTP Client**: WinHttp.WinHttpRequest.5.1
- **API Format**: OpenAI-compatible chat completions endpoint
- **Configuration Storage**: Windows Registry (HKCU\Software\OllamaAI)
- **UI Generation**: PowerShell script (invoked from VBA)
- **No external dependencies** — works with built-in Windows and Outlook components only

## Data Flow Diagram

```
Outlook (VBA)  --HTTP--> localhost:11434 -- Ollama Model
    ^                          |
    |                          v
    +------- JSON Response ----+
```

## Privacy

This addon is fully local:
- **No data** is sent to any external server
- **No API keys** are stored or transmitted
- **No telemetry**, analytics, or tracking
- **No internet connection required** after Ollama is installed
- All email processing happens on your machine

## Acknowledgments

Based on [key-assist-outlook](https://github.com/paulchi-intel/key-assist-outlook) by paulchi-intel. Recoded to replace Intel's proprietary ExpertGPT API with the open Ollama local inference server.

## License

MIT License

Copyright (c) 2026 cdblue999

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
