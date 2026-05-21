# Ollama Outlook AI — User Manual

## 1. Installation (first time)

**Step 1: Ensure Ollama is ready**
- Ollama must be installed and running (`http://localhost:11434`)
- Open terminal and run: `ollama pull llama3.2:3b`
- Verify with: `ollama list`

**Step 2: Import the macro into Outlook**
1. Open Outlook
2. Press **Alt+F11** to open the VBA editor
3. Go to **File > Import File** and select `ollama-outlook-ai.bas`
4. Close the VBA editor

**Step 3: Enable macros**
- Outlook 2013: **File > Options > Trust Center > Trust Center Settings > Macro Settings**
- Select: **Enable all macros**

**Step 4: Initialize (one time)**
1. Restart Outlook
2. Go to **View > Macros > View Macros**
3. Find and select `Ollama_Initialize`, click **Run**
4. Restart Outlook again

The **Ollama AI** button will now appear in email compose windows.

---

## 2. Daily Operation

### Using the AI assistant

1. **Compose a new email** or **Reply** to an existing one
2. Write the email body (the AI works with what you've written)
3. In the toolbar, click **Ollama AI > Process with AI**
4. Wait a few seconds — the AI response is inserted at your cursor position

### Changing settings

1. In any email compose window, click **Ollama AI > Settings**
2. A configuration window opens with three options:

| Setting | What it does |
|---------|-------------|
| **AI Model** | Select which Ollama model to use. Dropdown auto-populates from your installed models. |
| **System Prompt** | Instructions for the AI. Edit to change behavior (e.g., "Reply in Polish", "Summarize briefly"). |
| **Timeout** | How long to wait for a response. Increase for slower/larger models. |

3. Click **Save** to apply.

---

## 3. Troubleshooting

| Symptom | Fix |
|---------|-----|
| **"Ollama is not running"** | Start Ollama from Start Menu, or run `ollama serve` in terminal. |
| **No models in dropdown** | Open terminal: `ollama pull llama3.2:3b` then reopen Settings. |
| **AI takes too long** | Use a smaller model (3B-7B range) or increase timeout in Settings. |
| **Button doesn't appear** | Run `Ollama_Initialize` macro again, restart Outlook. |
| **Response is empty** | Check your model is downloaded (`ollama list`). Try a different model. |
| **VBA security warning** | Enable macros in Trust Center settings. |

---

## 4. Technical overview (for curious users)

```
Email text → VBA macro → HTTP POST → Ollama (localhost:11434) → AI model → Response → Inserted into email
```

- All traffic is **localhost only** — no internet needed
- Communication uses **OpenAI-compatible API format**
- Settings stored in **Windows Registry** (`HKCU\Software\OllamaAI`)
- Uses **WinHttp** for HTTP calls (built into Windows)
- No external DLLs, no additional software required

---

## 5. Files in this repository

| File | Purpose |
|------|---------|
| `ollama-outlook-ai.bas` | Main VBA module — import into Outlook |
| `README.md` | Project overview and installation guide |
| `MANUAL.md` | This file — daily operation reference |
