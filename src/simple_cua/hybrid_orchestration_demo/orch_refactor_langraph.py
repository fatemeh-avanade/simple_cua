"""
Vision-First Hybrid Orchestrator (LangGraph)
==============================================

Same capabilities as orch_refactor_stable_3.py but using LangGraph framework:
- State graph manages world state and orchestration
- Tools are LangChain tools with native image support
- Agent decides NEXT action (task-agnostic)
- Supports multimodal input (images) naturally
"""

import os
import time
import json
import base64
import subprocess
import re
from dataclasses import dataclass, field
from typing import Optional, Dict, Any, List, TypedDict, Annotated

import pyautogui
import pygetwindow as gw
from PIL import ImageGrab
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

from langchain_openai import AzureChatOpenAI
from langchain_core.messages import HumanMessage, BaseMessage, ToolMessage
from langchain_core.tools import tool
from langchain.tools import Tool
from langgraph.graph import StateGraph, END
from langgraph.prebuilt import ToolNode


# ---------------------------------------------------------------------
# ENV
# ---------------------------------------------------------------------

load_dotenv()

AZURE_ENDPOINT = "https://fa-test-openai-instance-canada-east.openai.azure.com/"
DEPLOYMENT_NAME = "fa-test-gpt-4o"

EDGE_PROFILE = r"C:\Users\fatemeh.torabi.asr\AppData\Local\Microsoft\Edge\User Data\Work"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(BASE_DIR, "data")
CMD_SCREENSHOT_PATH = os.path.join(DATA_PATH, "cmd.png")
EXCEL_SCREENSHOT_PATH = os.path.join(DATA_PATH, "excel.png")

os.makedirs(DATA_PATH, exist_ok=True)


# Initialize LLM
llm = AzureChatOpenAI(
    deployment_name=DEPLOYMENT_NAME,
    endpoint=AZURE_ENDPOINT,
    api_key=os.getenv("OPENAI_API_KEY"),
    api_version="2024-02-15-preview",
)


# ---------------------------------------------------------------------
# UTILS (generic)
# ---------------------------------------------------------------------

def safe_json_parse(text: str) -> dict:
    if not text or not text.strip():
        raise ValueError("Empty LLM response")

    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*", "", text)
        text = text.replace("```", "").strip()

    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise ValueError(f"No JSON found:\n{text}")

    return json.loads(match.group(0))


def human_escalation(state: dict, reason: str) -> None:
    print("\n⚠️ HUMAN ESCALATION REQUIRED")
    print("Reason:", reason)
    print("State:")
    print(json.dumps(state, indent=2, default=str))
    raise RuntimeError(reason)


# ---------------------------------------------------------------------
# STATE DEFINITION
# ---------------------------------------------------------------------

class GraphState(TypedDict):
    """State for the orchestration graph"""
    goal: Optional[Dict[str, Any]]
    
    # CMD
    cmd_rendered: bool
    cmd_image: Optional[str]
    cmd_value: Optional[float]
    
    # Excel
    excel_value: Optional[float]
    excel_attempts: List[Dict[str, Any]]
    
    # Result
    comparison: Optional[Dict[str, Any]]
    
    # Episodic memory
    history: List[Dict[str, Any]]
    
    # Messages for agent
    messages: List[BaseMessage]


# ---------------------------------------------------------------------
# TOOLS (LangChain @tool decorator)
# ---------------------------------------------------------------------

@tool
def parse_goal(prompt: str) -> str:
    """
    Parse user prompt into task parameters.
    Returns JSON with cmd, excel, and comparison_spec.
    """
    messages = [
        {
            "role": "system",
            "content": (
                "Extract task parameters.\n"
                "Return ONLY JSON with EXACT schema:\n"
                "{\n"
                '  "cmd": { "label": string, "file_path": string },\n'
                '  "excel": { "label": string, "url": string },\n'
                '  "comparison_spec": { "tolerance": number }\n'
                "}\n"
                "Labels must be literal tokens only (e.g., TOTAL, FV)."
            ),
        },
        {"role": "user", "content": prompt},
    ]

    response = llm.invoke(
        [HumanMessage(content=m["content"]) for m in messages]
    )

    parsed = safe_json_parse(response.content)
    parsed["cmd"]["label"] = parsed["cmd"]["label"].strip()
    parsed["excel"]["label"] = parsed["excel"]["label"].strip()
    return json.dumps(parsed)


@tool
def normalize_excel_url(url: str) -> str:
    """Ensure Excel opens in web viewer mode."""
    if "web=1" not in url:
        sep = "&" if "?" in url else "?"
        return f"{url}{sep}web=1"
    return url


@tool
def render_file_in_cmd(file_path: str) -> str:
    """Open CMD and display file contents."""
    subprocess.Popen("start cmd.exe", shell=True)
    time.sleep(2)

    wins = [
        w for w in gw.getAllWindows()
        if "Command Prompt" in w.title or "cmd" in w.title.lower()
    ]
    if not wins:
        raise RuntimeError("CMD window not found")

    win = wins[0]
    win.activate()
    time.sleep(0.5)

    pyautogui.write(f'type "{file_path}"', interval=0.03)
    pyautogui.press("enter")
    time.sleep(1.5)

    return "File rendered in CMD"


@tool
def capture_cmd() -> str:
    """Screenshot the CMD window and save to disk."""
    wins = [
        w for w in gw.getAllWindows()
        if "Command Prompt" in w.title or "cmd" in w.title.lower()
    ]
    if not wins:
        raise RuntimeError("CMD window not found")

    win = wins[0]
    bbox = (win.left, win.top, win.right, win.bottom)

    img = ImageGrab.grab(bbox=bbox)
    img.save(CMD_SCREENSHOT_PATH)

    print(f">>> CMD screenshot saved to {CMD_SCREENSHOT_PATH}")
    return CMD_SCREENSHOT_PATH


@tool
def extract_numeric_value_near_label(image_path: str, label: str) -> str:
    """
    Use Azure vision to extract numeric value from an image.
    Supports base64-encoded images or file paths.
    Returns JSON: {"value": number}
    """
    with open(image_path, "rb") as f:
        img_b64 = base64.b64encode(f.read()).decode("utf-8")

    # Use LangChain's native support for image URLs
    message = HumanMessage(
        content=[
            {"type": "text", "text": f"Extract numeric value near label '{label}'. Return JSON: {{ \"value\": number }}"},
            {
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{img_b64}"},
            },
        ]
    )

    response = llm.invoke([message])
    parsed = safe_json_parse(response.content)
    value = float(parsed["value"])
    print(f">>> Vision extracted value for '{label}': {value}")
    return json.dumps({"value": value})


@tool
def extract_excel_value_dom(excel_url: str, label: str) -> str:
    """
    Extract numeric value from Excel using DOM-based approach.
    Returns JSON result with status, value, reason, debug.
    """
    context = None
    page = None
    try:
        print(">>> [Excel:dom] launching...")

        with sync_playwright() as p:
            context, page = _open_excel_page(p, excel_url)

            page.screenshot(path=EXCEL_SCREENSHOT_PATH, full_page=True)
            try:
                size = os.path.getsize(EXCEL_SCREENSHOT_PATH)
                print(f">>> [Excel:dom] screenshot saved: {EXCEL_SCREENSHOT_PATH} ({size} bytes)")
            except Exception:
                print(f">>> [Excel:dom] screenshot saved: {EXCEL_SCREENSHOT_PATH}")

            probe = page.evaluate(
                """
                (label) => {
                    const elems = document.querySelectorAll('div, span, td, [role="gridcell"]');

                    for (const e of elems) {
                        const t = (e.innerText || e.textContent || '').trim();
                        if (t === label) {
                            const sib = e.nextElementSibling;
                            const v = sib ? (sib.innerText || sib.textContent || '').trim() : null;
                            return { found: true, value: v };
                        }
                    }

                    const all = document.querySelectorAll('*');
                    const sample = [];
                    for (let i = 0; i < Math.min(80, all.length); i++) {
                        const t = (all[i].innerText || all[i].textContent || '').trim();
                        if (t && t.length < 120) sample.push(t);
                    }
                    return { found: false, sample };
                }
                """,
                label
            )

            print(f">>> [Excel:dom] probe result for '{label}':", probe)

            if not isinstance(probe, dict) or probe.get("found") is False:
                return json.dumps({
                    "status": "not_found",
                    "value": None,
                    "reason": "DOM label not found (iframe/canvas/virtualized grid likely)",
                    "debug": probe.get("sample") if isinstance(probe, dict) else probe
                })

            raw = probe.get("value")
            if raw is None or str(raw).strip() == "":
                return json.dumps({
                    "status": "not_found",
                    "value": None,
                    "reason": "Label found but adjacent value missing/empty",
                    "debug": probe
                })

            try:
                return json.dumps({"status": "ok", "value": float(raw), "reason": None, "debug": None})
            except ValueError:
                return json.dumps({
                    "status": "error",
                    "value": None,
                    "reason": f"Adjacent value not numeric: {repr(raw)}",
                    "debug": probe
                })

    except Exception as e:
        print(">>> [Excel:dom] EXCEPTION:", repr(e))
        return json.dumps({"status": "error", "value": None, "reason": str(e), "debug": None})

    finally:
        try:
            if page:
                page.close()
        except Exception:
            pass
        try:
            if context:
                context.close()
        except Exception:
            pass


@tool
def extract_excel_value_vision(excel_url: str, label: str) -> str:
    """
    Extract numeric value from Excel using vision-based approach.
    Returns JSON result with status, value, reason, debug.
    """
    context = None
    page = None
    try:
        print(">>> [Excel:vision] launching...")

        with sync_playwright() as p:
            context, page = _open_excel_page(p, excel_url)

            page.screenshot(path=EXCEL_SCREENSHOT_PATH, full_page=True)
            try:
                size = os.path.getsize(EXCEL_SCREENSHOT_PATH)
                print(f">>> [Excel:vision] screenshot saved: {EXCEL_SCREENSHOT_PATH} ({size} bytes)")
            except Exception:
                print(f">>> [Excel:vision] screenshot saved: {EXCEL_SCREENSHOT_PATH}")

            # Use vision tool to extract from screenshot
            result_json = extract_numeric_value_near_label(EXCEL_SCREENSHOT_PATH, label)
            result = json.loads(result_json)
            val = result.get("value")
            print(f">>> [Excel:vision] extracted value for '{label}': {val}")

            return json.dumps({"status": "ok", "value": val, "reason": None, "debug": None})

    except Exception as e:
        print(">>> [Excel:vision] EXCEPTION:", repr(e))
        return json.dumps({"status": "error", "value": None, "reason": str(e), "debug": None})

    finally:
        try:
            if page:
                page.close()
        except Exception:
            pass
        try:
            if context:
                context.close()
        except Exception:
            pass


def _open_excel_page(p, excel_url: str):
    """Internal helper: open Excel Online workbook."""
    context = p.chromium.launch_persistent_context(
        EDGE_PROFILE,
        channel="msedge",
        headless=False,
        args=["--no-first-run", "--disable-extensions"],
    )
    page = context.new_page()

    page.set_default_timeout(120_000)

    page.goto(excel_url, wait_until="load", timeout=120_000)
    page.wait_for_timeout(12_000)

    try:
        page.wait_for_selector("canvas, [role='grid'], [role='gridcell']", timeout=15_000)
        print(">>> [Excel] readiness selector seen")
    except Exception:
        print(">>> [Excel] readiness selector NOT seen (continuing anyway)")

    try:
        print(">>> [Excel] page.url:", page.url)
        print(">>> [Excel] page.title:", page.title())
    except Exception:
        pass

    return context, page


# Collect all tools
tools = [
    parse_goal,
    normalize_excel_url,
    render_file_in_cmd,
    capture_cmd,
    extract_numeric_value_near_label,
    extract_excel_value_dom,
    extract_excel_value_vision,
]


# ---------------------------------------------------------------------
# ORCHESTRATOR AGENT (DECISION NODE)
# ---------------------------------------------------------------------

ACTIONS = [
    "render_cmd_file",
    "capture_cmd",
    "extract_cmd_value",
    "extract_excel_dom",
    "extract_excel_vision",
    "compare",
    "finish",
    "escalate",
]


def orchestrator_agent(state: GraphState) -> GraphState:
    """
    LLM-based agent that decides the next action.
    """
    print("\n>>> [Orchestrator Agent] deciding next action...")

    decision_prompt = json.dumps(
        {
            "goal": state["goal"],
            "world_state": {
                "cmd_rendered": state["cmd_rendered"],
                "cmd_image": state["cmd_image"],
                "cmd_value": state["cmd_value"],
                "excel_value": state["excel_value"],
                "excel_attempts": state["excel_attempts"],
                "comparison": state["comparison"],
            },
            "recent_steps": state["history"][-4:],
            "available_actions": ACTIONS,
        },
        indent=2,
        default=str,
    )

    system_message = (
        "You are an Orchestrator Agent.\n"
        "Decide the NEXT action only.\n"
        "Do not repeat actions that would not change the world.\n"
        "If stuck or uncertain, choose escalate.\n"
        "Return JSON: { \"action\": string }"
    )

    messages = [
        HumanMessage(content=system_message + "\n\n" + decision_prompt)
    ]

    print(">>> [Orchestrator Agent] messages:\n", messages)

    response = llm.invoke(messages)
    print(">>> [Orchestrator Agent] raw response:", response.content)

    decision = safe_json_parse(response.content)
    action = decision.get("action", "escalate")

    # Update history
    if not state["history"] or state["history"][-1]["action"] != action:
        state["history"].append({"action": action})

    state["messages"] = messages + [response]
    return state


# ---------------------------------------------------------------------
# ACTION NODES
# ---------------------------------------------------------------------

def action_render_cmd_file(state: GraphState) -> GraphState:
    """Execute render_cmd_file"""
    render_file_in_cmd(state["goal"]["cmd"]["file_path"])
    state["cmd_rendered"] = True
    return state


def action_capture_cmd(state: GraphState) -> GraphState:
    """Execute capture_cmd"""
    state["cmd_image"] = capture_cmd()
    return state


def action_extract_cmd_value(state: GraphState) -> GraphState:
    """Execute extract_cmd_value"""
    result_json = extract_numeric_value_near_label(
        state["cmd_image"], state["goal"]["cmd"]["label"]
    )
    result = json.loads(result_json)
    state["cmd_value"] = float(result["value"])
    return state


def action_extract_excel_dom(state: GraphState) -> GraphState:
    """Execute extract_excel_dom"""
    res = json.loads(extract_excel_value_dom(
        state["goal"]["excel"]["url"],
        state["goal"]["excel"]["label"]
    ))
    state["excel_attempts"].append({"mode": "dom", "result": res})
    if res["status"] == "ok":
        state["excel_value"] = res["value"]
    return state


def action_extract_excel_vision(state: GraphState) -> GraphState:
    """Execute extract_excel_vision"""
    res = json.loads(extract_excel_value_vision(
        state["goal"]["excel"]["url"],
        state["goal"]["excel"]["label"]
    ))
    state["excel_attempts"].append({"mode": "vision", "result": res})
    if res["status"] == "ok":
        state["excel_value"] = res["value"]
    return state


def action_compare(state: GraphState) -> GraphState:
    """Execute compare"""
    tol = state["goal"]["comparison_spec"]["tolerance"]
    state["comparison"] = {
        "match": abs(state["cmd_value"] - state["excel_value"]) <= tol,
        "cmd_value": state["cmd_value"],
        "excel_value": state["excel_value"],
    }
    return state


def action_finish(state: GraphState) -> GraphState:
    """Execute finish"""
    print("\n✅ FINAL RESULT")
    print(json.dumps(state["comparison"], indent=2))
    return state


def action_escalate(state: GraphState) -> GraphState:
    """Execute escalate"""
    human_escalation(state, "Orchestrator escalated")
    return state


# Routing function: map action to node
def route_action(state: GraphState) -> str:
    """Route to the appropriate action node based on last decision"""
    if not state["history"]:
        return "escalate"

    last_action = state["history"][-1].get("action", "escalate")

    if last_action == "render_cmd_file":
        return "action_render_cmd_file"
    elif last_action == "capture_cmd":
        return "action_capture_cmd"
    elif last_action == "extract_cmd_value":
        return "action_extract_cmd_value"
    elif last_action == "extract_excel_dom":
        return "action_extract_excel_dom"
    elif last_action == "extract_excel_vision":
        return "action_extract_excel_vision"
    elif last_action == "compare":
        return "action_compare"
    elif last_action == "finish":
        return "action_finish"
    elif last_action == "escalate":
        return "action_escalate"
    else:
        return "action_escalate"


# Should we continue looping?
def should_continue(state: GraphState) -> str:
    """Decide whether to continue or stop"""
    if not state["history"]:
        return "agent"

    last_action = state["history"][-1].get("action", "escalate")

    if last_action in ["finish", "escalate"]:
        return END
    else:
        return "agent"


# ---------------------------------------------------------------------
# BUILD GRAPH
# ---------------------------------------------------------------------

graph = StateGraph(GraphState)

# Add nodes
graph.add_node("agent", orchestrator_agent)
graph.add_node("action_render_cmd_file", action_render_cmd_file)
graph.add_node("action_capture_cmd", action_capture_cmd)
graph.add_node("action_extract_cmd_value", action_extract_cmd_value)
graph.add_node("action_extract_excel_dom", action_extract_excel_dom)
graph.add_node("action_extract_excel_vision", action_extract_excel_vision)
graph.add_node("action_compare", action_compare)
graph.add_node("action_finish", action_finish)
graph.add_node("action_escalate", action_escalate)

# Add edges
graph.set_entry_point("agent")
graph.add_conditional_edges(
    "agent",
    route_action,
    {
        "action_render_cmd_file": "action_render_cmd_file",
        "action_capture_cmd": "action_capture_cmd",
        "action_extract_cmd_value": "action_extract_cmd_value",
        "action_extract_excel_dom": "action_extract_excel_dom",
        "action_extract_excel_vision": "action_extract_excel_vision",
        "action_compare": "action_compare",
        "action_finish": "action_finish",
        "action_escalate": "action_escalate",
    }
)

# All action nodes route back to agent (except finish/escalate which go to END)
for action_node in [
    "action_render_cmd_file",
    "action_capture_cmd",
    "action_extract_cmd_value",
    "action_extract_excel_dom",
    "action_extract_excel_vision",
    "action_compare",
]:
    graph.add_edge(action_node, "agent")

graph.add_edge("action_finish", END)
graph.add_edge("action_escalate", END)

runnable = graph.compile()


# ---------------------------------------------------------------------
# RUN
# ---------------------------------------------------------------------

def run(prompt: str):
    """Main orchestration entry point"""
    # Parse goal
    goal_json = parse_goal(prompt)
    goal = json.loads(goal_json)

    # Normalize Excel URL
    goal["excel"]["url"] = normalize_excel_url(goal["excel"]["url"])

    print("\n>>> Parsed goal:\n", json.dumps(goal, indent=2))

    # Initialize state
    initial_state = GraphState(
        goal=goal,
        cmd_rendered=False,
        cmd_image=None,
        cmd_value=None,
        excel_value=None,
        excel_attempts=[],
        comparison=None,
        history=[],
        messages=[],
    )

    # Run graph
    final_state = runnable.invoke(initial_state)

    return final_state


# ---------------------------------------------------------------------

if __name__ == "__main__":
    user_prompt = r"""
Find TOTAL in:
"C:\Users\fatemeh.torabi.asr\small_vsc_projects\quick_python_project\quick_python_project\hybrid_orchestration_demo\data\invoice.txt"
and FV in the spreadsheet at
https://avanade-my.sharepoint.com/:x:/r/personal/fatemeh_torabi_asr_avanade_com/Documents/test_data_folder/Book.xlsx
Compare with tolerance 0.01.
"""
    result = run(user_prompt)
    print("\n>>> Final state comparison:", result["comparison"])
