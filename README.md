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

## Test Run Results

All versions tested against the same task: extract `TOTAL` from a local text invoice via CMD screenshot + vision, extract `FV` from an Excel Online workbook, compare within tolerance 0.01.

**Verified output (all passing versions):**
```json
{ "match": true, "cmd_value": 1240.5, "excel_value": 1240.5 }
```

### Summary table

| Version | Result | Brain calls | Duration | Notes |
|---|---|---|---|---|
| `orchestration_example` (V0) | ❌ Crashed | 1 | <1s | Naive JSON parser crashes on markdown-wrapped LLM response |
| `orch_refactor_stable` (V1) | ✅ Pass | 6 | ~2-3 min | DOM attempted first (failed — Excel Online uses canvas), pivoted to vision |
| `orch_refactor_stable_2` (V2) | ✅ Pass | 5 | ~2-3 min | Policy brain skipped DOM entirely — 1 fewer LLM call than V1 |
| `orch_refactor_stable_3` (V3) | ✅ Pass | 6 | ~2-3 min | Task-agnostic world state; per-step `log_world()` output |
| `orch_refactor_langraph` | ✅ Pass | 6 | ~5-6 min | Needs fixes before running (see below) |
| `orch_refactor_semantic_kernel` | ✅ Pass | 7 | ~5-6 min | Needs fixes before running (see below) |
| `orch_refactor_wip.py` | — | — | — | Work-in-progress scratch file — not a runnable orchestrator |

---

### V0 — Why it crashed

`orchestration_example.py` calls `json.loads()` directly on the raw LLM response string. GPT-4o frequently wraps its output in markdown fences (` ```json ... ``` `) or returns explanatory prose around the JSON. With no stripping or regex fallback, the first `json.loads()` raises a `JSONDecodeError` and the script exits immediately. All later versions fix this with `safe_json_parse()`, which strips markdown fences and uses a regex to extract the JSON object.

---

### V1/V3 — Why DOM always fails on Excel Online

Excel Online renders spreadsheet content using a `<canvas>` element and a virtualized grid — there are no actual `<td>` or `<span>` elements in the DOM containing cell values. The DOM probe in `extract_excel_value_dom` queries `div, span, td, [role="gridcell"]` but finds nothing matching the label. V1 and V3 correctly detect the `not_found` status and the brain pivots to the vision fallback. V2's policy brain learns from the history context that DOM failed previously and skips straight to vision on the next attempt.

---

### V2 — Why it uses one fewer brain call

V2's brain is a pure policy agent: instead of the orchestrator hardcoding a "try DOM first, then vision" rule, the LLM reasons from recent history. By the time it needs to extract the Excel value, it has already seen that DOM extraction in the current session produced a `not_found` result. It therefore decides to go straight to `extract_excel_vision`, skipping the redundant DOM attempt and saving one full browser launch cycle (~30–60s of network + render time) plus one LLM call.

---

### V3 — Why it uses 6 calls despite task-agnostic state

V3 uses the same step sequence as V1 (render → capture → extract CMD → try DOM → try vision → compare → finish). The task-agnostic `WorldState` doesn't change the number of decisions the brain makes — it just makes the world model extensible. The 6th call is `finish` (the brain must explicitly decide the run is complete).

---

### LangGraph port

Uses LangChain's `AzureChatOpenAI`, tools decorated with `@tool` and registered with `ToolNode`, and a compiled `StateGraph` replacing the `while True` loop. Output is fully buffered by `runnable.invoke()` — nothing appears until the graph completes, which is why it feels slower than V1–V3 that print incrementally.

---

### Semantic Kernel port

Uses SK's `Kernel` + `AzureChatCompletion`, tools decorated with `@kernel_function` registered as plugins, and `async_playwright` for browser automation (required because the run loop is an `asyncio` coroutine). The extra brain call (7 vs 6) comes from SK's async dispatcher adding one scheduling round-trip per plugin invocation.

---

### Why LangGraph and Semantic Kernel are slower

The actual task work (browser automation, vision API round-trip) takes roughly the same time across all versions. The extra time in the framework ports comes from:

**LangGraph:**
- **Pregel runtime**: every node activation goes through the graph scheduler and task queue
- **LangChain callback pipeline**: `@tool` calls fire tracing/callback hooks on every invocation, even when no callbacks are registered
- **Buffered output**: `runnable.invoke()` holds all stdout until the full graph completes — the run looks silent but is working

**Semantic Kernel:**
- **Async event loop scheduling**: all plugin calls go through SK's async dispatcher, adding overhead even for synchronous tools
- **Kernel service resolution**: `get_service(type=...)` traverses the kernel registry on every LLM call
- **Buffered output**: same `asyncio.run()` buffering pattern as LangGraph

For an interactive desktop agent (2-3 min task), this overhead is noticeable but not prohibitive. For high-frequency or latency-sensitive agents it would matter more.

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
