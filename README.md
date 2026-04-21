# Doc Editorial Review

> **AI-powered editorial review for PowerPoint decks.** Upload a `.pptx` (or a folder of decks) and get a structured QA report covering **spelling, grammar, punctuation, terminology, tone, visual consistency, and layout** — powered by **Microsoft Azure AI Foundry** agents.

[![Python 3.12+](https://img.shields.io/badge/python-3.12%2B-blue.svg)](https://www.python.org/downloads/)
[![Azure AI Foundry](https://img.shields.io/badge/Azure-AI%20Foundry-0078D4.svg)](https://ai.azure.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## ✨ Features

- **End-to-end PPTX review** — text content + visual/layout consistency in one pass
- **Multi-pass editorial reviewer** — inventory → per-slide QA → cross-slide consistency → verification (mandatory false-positive filter)
- **Deterministic visual analysis** — fonts, colors, sizes, positions, overlap, alignment (no LLM token cost)
- **Adaptive chunking** — single pass for small decks, sliding-window with overlap for 200+ slide decks
- **Cross-chunk context** via Foundry's `previous_response_id` — terminology drift caught across chunks
- **Local-first data flow** — `.pptx` files never leave your machine; only extracted text is sent to the model
- **Gradio web UI** — model selector, editable instructions, 3-tab results (summary, per-slide drill-down, full report)
- **Multi-deck batch mode** — review every `.pptx` in a folder in one click
- **Container-ready** — Dockerfile + Bicep template for Azure Container Apps deployment

---

## 🏗️ Architecture

```
┌──────────────────┐                ┌──────────────────────┐
│  Gradio Web UI   │ ────────────▶ │  Local Python (app)  │
│  localhost:7860  │ ◀──────────── │  • PPTX extraction   │
└──────────────────┘                │  • Visual analysis   │
                                    │  • Chunking + dedup  │
                                    │  • HTML rendering    │
                                    └──────────┬───────────┘
                                               │ Responses API
                                               │ (text only)
                                               ▼
                                    ┌──────────────────────┐
                                    │  Azure AI Foundry    │
                                    │  • Prompt agent      │
                                    │  • LLM reasoning     │
                                    │  • 6 FunctionTools   │
                                    └──────────────────────┘
```

**Split-brain by design**: All data processing (PPTX parsing, visual checks, dedup, rendering) happens locally. The Foundry agent only sees extracted text and returns editorial findings. PPTX files never leave your machine.

For a deep architecture dive, see [SOLUTION_OVERVIEW.md](SOLUTION_OVERVIEW.md).

---

## 📋 Prerequisites

- **Python 3.12+**
- **Azure CLI** ([install](https://learn.microsoft.com/cli/azure/install-azure-cli))
- **An Azure AI Foundry project** with at least one chat-model deployment (e.g., `gpt-4o`)
  - [Create a Foundry project](https://learn.microsoft.com/azure/ai-foundry/how-to/create-projects)
  - [Deploy a model](https://learn.microsoft.com/azure/ai-foundry/how-to/deploy-models-openai)

---

## 🚀 Quick Start

### 1. Clone the repo

```bash
git clone https://github.com/dhangerkapil/doc-editorial-review.git
cd doc-editorial-review
```

### 2. Create a virtual environment

```powershell
# Windows (PowerShell)
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

```bash
# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 3. Configure your Foundry project

Copy the example env file and fill in your values:

```bash
cp .env.example .env
```

Edit `.env`:

```ini
AZURE_AI_PROJECT_ENDPOINT=https://<your-resource>.services.ai.azure.com/api/projects/<your-project>
AZURE_AI_MODEL_DEPLOYMENT_NAME=gpt-4o
```

> 💡 You can find the project endpoint on the **Overview** page of your Foundry project. Deployment names are in the **Models + endpoints** tab.

### 4. Sign in to Azure

```bash
az login
# Optionally pin to a specific tenant:
az login --tenant <your-tenant-id>
```

### 5. Run the app

```bash
python app.py
```

Open **http://127.0.0.1:7860** in your browser.

---

## 🖥️ Using the App

The UI offers three ways to review:

| Method | Use when… |
|---|---|
| **📂 Select Folder** | You want to review every `.pptx` in a folder (recursive). Native folder picker. |
| **📁 Path Input** | You want to type/paste a folder path manually. |
| **📎 Upload PPTX** | You want to review a single deck via drag-and-drop. |

Results appear in three tabs:

- **📊 Summary** — severity counts, category breakdown, quick-look findings table
- **🔍 Per-Slide Findings** — expandable per-slide drill-down with evidence and suggested fixes
- **📝 Full Agent Report** — raw markdown report from the LLM (glossary, tone analysis, full findings)

The **Agent Instructions** panel is editable — tweak the system prompt without touching code.

---

## 🧪 Try the Notebook

A walkthrough notebook is included for SDK-level exploration:

```bash
jupyter lab editorial_qa_agent.ipynb
```

It covers: agent creation, function tool registration, the orchestration loop, and result rendering.

---

## 📁 Project Structure

```
doc-editorial-review/
├── app.py                      # Main Gradio web app + agent orchestration
├── editorial_qa_agent.ipynb    # SDK walkthrough notebook
├── requirements.txt            # Python dependencies
├── .env.example                # Template for your Foundry config
├── Dockerfile                  # Container image for Azure deployment
├── deploy.ps1                  # End-to-end Azure Container Apps deployment
├── infra/
│   └── main.bicep              # Azure infrastructure (ACR + Container App + RBAC)
├── SOLUTION_OVERVIEW.md        # Deep architecture & function reference
├── LICENSE                     # MIT
└── README.md                   # This file
```

---

## ☁️ Deploy to Azure Container Apps

A one-shot PowerShell script and Bicep template are included. From the repo root:

```powershell
$RG = "rg-editorial-qa"
$LOCATION = "eastus2"
$FOUNDRY_ENDPOINT = "https://your-resource.services.ai.azure.com/api/projects/your-project"

.\deploy.ps1 -ResourceGroup $RG -Location $LOCATION -FoundryEndpoint $FOUNDRY_ENDPOINT
```

The script will:

1. Create the resource group
2. Deploy infra via Bicep (Azure Container Registry + Container App + managed identity)
3. Build & push the Docker image
4. Update the Container App to the new image
5. Print the public URL

The Container App uses a **system-assigned managed identity** to authenticate to your Foundry project — no keys, no secrets.

---

## 🔧 How It Works (TL;DR)

The agent calls 6 deterministic Python functions (registered as `FunctionTool` schemas):

| Tool | Purpose |
|---|---|
| `extract_deck` | Parse PPTX → markdown with slide markers |
| `analyze_visual_consistency` | Compute dominant title/body patterns and flag deviations |
| `get_review_windows` | Single pass (<90K tokens) or sliding window (1200 words / 250 overlap) |
| `store_chunk_findings` | Append findings from each chunk to an accumulator |
| `merge_and_dedupe_findings` | Deduplicate by `(slide + issue + evidence)`, keep highest severity |
| `extract_deck_visual` | Raw visual metadata dump (fallback) |

For each text chunk the LLM runs **4 passes**: inventory (build canonical glossary + tone), per-slide QA, cross-slide consistency, and a mandatory verification pass that drops ≥30% of candidate findings to enforce a **zero false-positive policy**.

Cover slide (slide 1) and the closing slide are excluded from visual checks — they're expected to be styled differently.

---

## 🔒 Security & Privacy

- **No keys stored** — auth is via `AzureCliCredential` locally and managed identity in Azure
- **No data leaves your machine** other than extracted slide text sent to the model
- **No embeddings, no vector DB, no RAG** — the entire deck is processed end-to-end
- All HTML output is escaped to prevent XSS in the Gradio UI

---

## 🛠️ Troubleshooting

| Symptom | Fix |
|---|---|
| `AzureCliCredential.get_token_info failed` | Run `az login` (or `az login --tenant <id>`) |
| `Cannot find empty port in range: 7860-7860` | Another process is using the port. Stop it, or set `GRADIO_SERVER_PORT=7861` and pass `server_port` in `app.launch()` |
| `Could not list deployments` | Verify `AZURE_AI_PROJECT_ENDPOINT` in `.env`; your CLI identity must have `Cognitive Services User` role on the Foundry project |
| Empty findings | Try a different model (gpt-4o is a good baseline); confirm the deck has actual text content |

---

## 🤝 Contributing

Issues and pull requests are welcome. Please:

1. Open an issue describing the problem or proposal
2. Keep PRs focused and include a brief description of the change
3. Verify the app still launches (`python app.py`) before submitting

---

## 📄 License

[MIT](LICENSE) © 2026 Kapil Dhanger
