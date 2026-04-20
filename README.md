# simple-cua

A collection of experiments for building **hybrid LLM-brained computer use agents** — autonomous orchestrators that combine GPT-4o vision, browser automation (Playwright), terminal interaction (pyautogui), and structured world state to complete multi-step desktop tasks.

---

## Quick Start

### Prerequisites

- Python 3.13
- [Poetry](https://python-poetry.org/docs/#installation)
- Microsoft Edge (for Playwright persistent context)
- Azure OpenAI access (GPT-4o deployment)

### 1. Clone & install

```bash
git clone https://github.com/fatemeh-avanade/simple_cua.git
cd simple_cua
poetry install
```

### 2. Configure environment

Create a `.env` file in the project root:

```env
OPENAI_API_KEY=your_azure_openai_key
```

> All scripts read `OPENAI_API_KEY` via `python-dotenv`. The Azure endpoint and deployment name are set as constants near the top of each file — update them to match your Azure OpenAI instance.

### 3. Activate the environment

```bash
poetry shell
```

Or run a single script directly:

```bash
poetry run python src/simple_cua/hybrid_orchestration_demo/orch_refactor_stable_3.py
```

### 4. VS Code interpreter

The `.vscode/settings.json` already points to the Poetry virtualenv. Open the folder in VS Code and the correct Python 3.13 interpreter will be selected automatically.

---

## Project Structure

```
src/simple_cua/
├── hybrid_orchestration_demo/     # Orchestration experiments (see below)
│   ├── orchestration_example.py   # V0 — naive original
│   ├── orch_refactor_stable.py    # V1 — typed state + episodic memory
│   ├── orch_refactor_stable_2.py  # V2 — policy-level brain
│   ├── orch_refactor_stable_3.py  # V3 — clean architecture (recommended)
│   ├── orch_refactor_wip.py       # WIP scratch
│   ├── orch_refactor_langraph.py  # V3 ported to LangGraph
│   ├── orch_refactor_semantic_kernel.py  # V3 ported to Semantic Kernel
│   ├── browser_excel_test.py      # Standalone Playwright/Excel test
│   ├── cmd_grid_test.py           # Standalone CMD interaction test
│   ├── gpt_ocr_test.py            # Standalone GPT-4o OCR test
│   └── pyautogui_test.py          # Standalone pyautogui test
└── omni_ui_agent/                 # Separate vision agent experiments
    └── src/
        ├── agent_run.py
        ├── vision_extract.py
        └── hello_world.py
```

---

## Orchestration File Evolution

The `hybrid_orchestration_demo` folder contains a progressive series of orchestration implementations, each addressing the shortcomings of the previous version. The task throughout is the same: extract a numeric value from a local file via CMD screenshot + vision, extract another from Excel Online, and compare them within a tolerance.

---

### `orchestration_example.py` — V0: Naive Original

- **State**: Plain `dict` with hardcoded keys (`terminal`, `vision`, `excel`, `comparison`)
- **Memory**: None — only current snapshot, no history
- **Agent interaction**: Monolithic classes (`TerminalAgent`, `VisionAgent`, `ExcelAgent`) instantiated inside the run loop
- **Brain**: Sends the full state dict to GPT on every tick with no episodic context
- **Tools**: Stateful agent objects, not transactional functions
- **Limitations**: Hardcoded absolute paths, fragile JSON parsing, no retry logic, Excel has one strategy only (DOM)

---

### `orch_refactor_stable.py` — V1: First Refactor

- **State**: Typed `@dataclass State` with explicit fields + `errors` list
- **Memory**: Generic episodic `history` list introduced via `log_step()`
- **Brain**: Starts using `response_format={"type": "json_object"}` for reliable parsing; passes recent history window (last 6 steps)
- **Tools**: Flat functions replacing agent classes; CMD split into `render_file_in_cmd` + `capture_cmd_screenshot`
- **Excel**: Two independent transactional strategies — `extract_excel_value_dom` and `extract_excel_value_vision`
- **Improvement**: `safe_json_parse` uses regex to robustly extract JSON from messy LLM output

---

### `orch_refactor_stable_2.py` — V2: Episodic Memory + Policy Brain

- **State**: Typed dataclass, but domain-specific flags replaced by pure episodic `history`
- **Memory**: Brain receives only `recent_history` + a lean `current_state` snapshot — not the whole world
- **Brain**: Promoted to a **policy-level agent** — no hardcoded step rules in the system prompt. It reasons purely from history and what is still missing
- **Orchestrator**: Sends `goal + recent_history + current_state` (not full state) — leaner, less prompt bloat
- **Notable**: URL normalization added as a task-layer concern, separate from tools

---

### `orch_refactor_stable_3.py` — V3: Clean Architecture ✅ Recommended

- **State**: Fully **task-agnostic** `WorldState` — no domain fields, just `values: dict`, `artifacts: dict`, `history`, `errors`, `done`
- **Separation of concerns**:
  - **Orchestrator**: knows nothing about CMD/Excel/vision — only sees `known_values`, `known_artifacts`, `recent_history`
  - **Task Agents** (`compare_task_agent`, `copy_paste_task_agent`): pluggable callables that consume and produce `world.values`
  - **Tools**: pure transactional functions, no shared state
- **Multi-task support**: First version to support multiple task types with a single run loop — switch task by passing a different agent callable
- **Error handling**: Exceptions caught in run loop → appended to `world.errors` → orchestrator decides whether to retry or escalate
- **Debug**: `log_world()` prints a compact per-step snapshot of values, artifacts, and history

---

### `orch_refactor_langraph.py` — LangGraph Port of V3

- Replaces raw `AzureOpenAI` client with `AzureChatOpenAI` (LangChain)
- Tools decorated with `@tool` and registered with LangGraph's `ToolNode`
- World state implemented as a `TypedDict` for LangGraph's `StateGraph`
- Message passing uses native `HumanMessage` / `ToolMessage` types
- The `while True` orchestration loop becomes a **compiled state graph** with explicit edges
- **Benefit**: Checkpointing, branching, streaming, and observability are framework-native

---

### `orch_refactor_semantic_kernel.py` — Semantic Kernel Port of V3

- Replaces `AzureOpenAI` client with SK's `Kernel` + `AzureChatCompletion`
- Tools decorated with `@kernel_function` and registered as SK plugins
- Brain calls use `KernelArguments` for structured parameter passing; LLM calls go through `kernel.invoke()`
- Run loop structure is identical to V3
- **Benefit**: Pluggable AI service backends; SK memory and planner integration possible

---

### Evolution Summary

| File | State Model | Memory | Brain Style | Multi-task | Framework |
|---|---|---|---|---|---|
| `orchestration_example` | plain dict | ❌ | full-state dump | ❌ | raw OpenAI |
| `orch_refactor_stable` | typed dataclass | episodic log | full-state + history | ❌ | raw OpenAI |
| `orch_refactor_stable_2` | typed dataclass | episodic only | policy-level | ❌ | raw OpenAI |
| `orch_refactor_stable_3` | task-agnostic dict | episodic only | policy-level | ✅ | raw OpenAI |
| `orch_refactor_langraph` | TypedDict graph | message history | graph edges | ✅ | LangGraph |
| `orch_refactor_semantic_kernel` | task-agnostic dict | episodic | policy-level | ✅ | Semantic Kernel |

The key architectural leaps:
- **V1 → V2**: Brain becomes policy-level (reasons from history) rather than procedural (follows steps)
- **V2 → V3**: World state becomes fully task-agnostic, enabling pluggable task agents and multi-task support

---

## Dependencies

| Package | Purpose |
|---|---|
| `openai` | Azure OpenAI API client |
| `python-dotenv` | `.env` file loading |
| `playwright` | Browser automation (Excel Online) |
| `pyautogui` | Desktop GUI automation (CMD) |
| `pygetwindow` | Window management |
| `pillow` | Screenshot capture |
| `pandas` | Data utilities |
| `langchain` / `langchain-openai` / `langgraph` | LangGraph orchestration variant |
| `semantic-kernel` | Semantic Kernel orchestration variant |

---

## Notes

- All scripts target **Azure OpenAI** (GPT-4o). Update `AZURE_ENDPOINT` and `DEPLOYMENT_NAME` constants per file if your deployment differs.
- Excel Online automation uses a persistent Edge profile (`EDGE_PROFILE`). Update this path to your local Edge profile.
- Data files (screenshots, invoices) are written to a `data/` folder inside each module — this folder is gitignored.
- `semantic-kernel` requires Python `<3.14` due to its `openapi-core` dependency. This project targets Python 3.13.
