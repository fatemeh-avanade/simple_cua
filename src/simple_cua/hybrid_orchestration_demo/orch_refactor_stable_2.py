"""
Vision-First Hybrid Orchestrator (Refactored) — Episodic Memory Variant
======================================================================

Goal:
- Extract numeric value near a label from a local text file rendered in CMD (vision-only)
- Extract numeric value near a label from an Excel Online workbook using TWO independent strategies:
    1) DOM-based (best effort; often fails for Excel Online grids)
    2) Vision-based (screenshot + vision extraction)
- Compare the two values within tolerance
- Brain (LLM) chooses the next action at runtime (DOM vs Vision, retries, etc.)
- Tools are transactional (no persistent Excel handler context across tool calls)

NEW STRATEGY:
- Replace domain-specific memory (excel_attempts / dom_tried / vision_tried) with generic episodic history.
- Make the system prompt policy-level (underspecified) and provide memory as recent trajectory.
"""

import os
import time
import json
import base64
import subprocess
from dataclasses import dataclass, field
from typing import Optional, Dict, Any
import re

import pyautogui
import pygetwindow as gw
from PIL import ImageGrab
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv
from openai import AzureOpenAI


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

client = AzureOpenAI(
    api_key=os.getenv("OPENAI_API_KEY"),
    api_version="2024-02-15-preview",
    azure_endpoint=AZURE_ENDPOINT
)


# ---------------------------------------------------------------------
# UTILS
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


def log_step(state: "State", action: str, result: Any = None) -> None:
    """
    Generic episodic memory logger.
    Stores a compact trajectory of what happened, without domain-specific flags.
    """
    state.history.append({
        "step": len(state.history) + 1,
        "action": action,
        "result": result,
        "snapshot": {
            "cmd_rendered": state.cmd_rendered,
            "cmd_image_present": state.cmd_image is not None,
            "cmd_value": state.cmd_value,
            "excel_value": state.excel_value,
            "comparison_present": state.comparison is not None,
        },
    })


# ---------------------------------------------------------------------
# STATE
# ---------------------------------------------------------------------

@dataclass
class State:
    goal: Optional[Dict[str, Any]] = None

    # CMD
    cmd_rendered: bool = False
    cmd_image: Optional[str] = None
    cmd_value: Optional[float] = None

    # Excel
    excel_value: Optional[float] = None

    # Final
    comparison: Optional[Dict[str, Any]] = None

    # Generic episodic memory (NEW)
    history: list = field(default_factory=list)

    errors: list = field(default_factory=list)


# ---------------------------------------------------------------------
# CMD TOOLS (render -> capture -> vision extract)
# ---------------------------------------------------------------------

def render_file_in_cmd(file_path: str) -> None:
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


def capture_cmd_screenshot() -> str:
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


# ---------------------------------------------------------------------
# VISION TOOL
# ---------------------------------------------------------------------

def extract_numeric_value_near_label(image_path: str, label: str) -> float:
    with open(image_path, "rb") as f:
        img_b64 = base64.b64encode(f.read()).decode("utf-8")

    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": "Return JSON: { \"value\": number }"},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": f"Extract numeric value near label '{label}'. Label is exact."},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_b64}"}},
                ],
            },
        ],
    )

    parsed = json.loads(response.choices[0].message.content)
    print(f">>> Vision extracted value for '{label}': {parsed['value']}")
    return float(parsed["value"])


# ---------------------------------------------------------------------
# EXCEL TOOLS (two independent transactional tools)
# ---------------------------------------------------------------------

def _open_excel_page(p, excel_url: str):
    """
    Internal helper: open Excel Online workbook and return (context, page).
    No context is persisted across tool calls; each tool invocation is transactional.
    """
    context = p.chromium.launch_persistent_context(
        EDGE_PROFILE,
        channel="msedge",
        headless=False,
        args=["--no-first-run", "--disable-extensions"],
    )
    page = context.new_page()

    # Excel Online is slow + chatty; do not use networkidle.
    page.set_default_timeout(120_000)

    page.goto(excel_url, wait_until="load", timeout=120_000)
    page.wait_for_timeout(12_000)  # settle time

    # Best-effort readiness probe (non-fatal)
    try:
        page.wait_for_selector("canvas, [role='grid'], [role='gridcell']", timeout=15_000)
        print(">>> [Excel] readiness selector seen")
    except Exception:
        print(">>> [Excel] readiness selector NOT seen (continuing anyway)")

    # Diagnostics
    try:
        print(">>> [Excel] page.url:", page.url)
        print(">>> [Excel] page.title:", page.title())
    except Exception:
        pass

    return context, page


def extract_excel_value_dom(excel_url: str, label: str) -> dict:
    """
    Transactional DOM-based read.
    Returns: {status, value, reason, debug}
    """
    context = None
    page = None
    try:
        print(">>> [Excel:dom] launching...")

        with sync_playwright() as p:
            context, page = _open_excel_page(p, excel_url)

            # Always screenshot for debugging
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

                    // Debug sample if not found
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
                return {
                    "status": "not_found",
                    "value": None,
                    "reason": "DOM label not found (iframe/canvas/virtualized grid likely)",
                    "debug": probe.get("sample") if isinstance(probe, dict) else probe
                }

            raw = probe.get("value")
            if raw is None or str(raw).strip() == "":
                return {
                    "status": "not_found",
                    "value": None,
                    "reason": "Label found but adjacent value missing/empty",
                    "debug": probe
                }

            try:
                return {"status": "ok", "value": float(raw), "reason": None, "debug": None}
            except ValueError:
                return {
                    "status": "error",
                    "value": None,
                    "reason": f"Adjacent value not numeric: {repr(raw)}",
                    "debug": probe
                }

    except Exception as e:
        print(">>> [Excel:dom] EXCEPTION:", repr(e))
        return {"status": "error", "value": None, "reason": str(e), "debug": None}

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


def extract_excel_value_vision(excel_url: str, label: str) -> dict:
    """
    Transactional vision-based read.
    Returns: {status, value, reason, debug}
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

            val = extract_numeric_value_near_label(EXCEL_SCREENSHOT_PATH, label)
            print(f">>> [Excel:vision] extracted value for '{label}': {val}")

            return {"status": "ok", "value": val, "reason": None, "debug": None}

    except Exception as e:
        print(">>> [Excel:vision] EXCEPTION:", repr(e))
        return {"status": "error", "value": None, "reason": str(e), "debug": None}

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


# ---------------------------------------------------------------------
# BRAIN (goal parsing + next-action selection)
# ---------------------------------------------------------------------

ALLOWED_ACTIONS = [
    "render_cmd_file",
    "capture_cmd",
    "extract_cmd_value",
    "extract_excel_dom",
    "extract_excel_vision",
    "compare",
    "finish",
    "escalate",
]


def parse_goal(prompt: str) -> dict:
    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": (
                    "Extract task parameters.\n"
                    "IMPORTANT: cmd.label and excel.label must be the literal label tokens only (e.g., TOTAL, FV). "
                    "Do NOT include surrounding sentence text.\n"
                    "Return JSON:\n"
                    "{\n"
                    "  \"cmd\": { \"label\": string, \"file_path\": string },\n"
                    "  \"excel\": { \"label\": string, \"url\": string },\n"
                    "  \"comparison_spec\": { \"tolerance\": number }\n"
                    "}"
                ),
            },
            {"role": "user", "content": prompt},
        ],
    )
    parsed = json.loads(response.choices[0].message.content)
    parsed["cmd"]["label"] = parsed["cmd"]["label"].strip()
    parsed["excel"]["label"] = parsed["excel"]["label"].strip()
    return parsed


def decide_next_action(state: State) -> str:
    """
    Policy-level LLM brain:
    - No procedural step rules.
    - Uses recent episodic history to infer what to do next.
    """

    MEMORY_WINDOW = 6
    recent_history = state.history[-MEMORY_WINDOW:]

    payload = {
        "goal": state.goal,
        "recent_history": recent_history,
        "current_state": {
            "cmd_value_present": state.cmd_value is not None,
            "excel_value_present": state.excel_value is not None,
            "comparison_present": state.comparison is not None,
        },
        "available_actions": ALLOWED_ACTIONS,
    }

    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": (
                    "You are an autonomous orchestration agent.\n\n"
                    "Your task is to complete the user's goal by deciding the NEXT action only.\n\n"
                    "You are NOT given explicit step-by-step instructions.\n"
                    "You must infer what to do next based on:\n"
                    "- the goal\n"
                    "- what has already happened\n"
                    "- what information is still missing\n\n"
                    "General principles:\n"
                    "- Prefer progress over repetition.\n"
                    "- Avoid repeating actions that already failed or produced no new information.\n"
                    "- Use alternative strategies when one approach does not work.\n"
                    "- Once a value is obtained, do not try to obtain it again.\n"
                    "- When the task is complete, finish.\n"
                    "- If the task cannot be completed reliably, escalate.\n\n"
                    "Choose exactly ONE next action.\n\n"
                    "Return JSON only:\n"
                    "{ \"action\": string, \"reason\": string }"
                ),
            },
            {
                "role": "user",
                "content": json.dumps(payload, indent=2, default=str),
            },
        ],
    )

    try:
        decision = json.loads(response.choices[0].message.content)
        action = decision.get("action")
        reason = decision.get("reason", "")
        print("\n🧠 Brain rationale:", reason)

        if action not in ALLOWED_ACTIONS:
            print(">>> Invalid action; escalating.")
            return "escalate"

        return action
    except Exception as e:
        print(">>> Failed to parse brain output; escalating:", repr(e))
        return "escalate"


# ---------------------------------------------------------------------
# ORCHESTRATOR LOOP
# ---------------------------------------------------------------------

def run(prompt: str):
    state = State(goal=parse_goal(prompt))
    print("\n>>> Parsed goal:\n", json.dumps(state.goal, indent=2))

    while True:
        action = decide_next_action(state)
        print(f"\n🧠 Brain decided: {action}")

        if action == "render_cmd_file":
            render_file_in_cmd(state.goal["cmd"]["file_path"])
            state.cmd_rendered = True
            log_step(state, action, "ok")

        elif action == "capture_cmd":
            state.cmd_image = capture_cmd_screenshot()
            log_step(state, action, {"image": state.cmd_image})

        elif action == "extract_cmd_value":
            state.cmd_value = extract_numeric_value_near_label(
                state.cmd_image,
                state.goal["cmd"]["label"]
            )
            log_step(state, action, state.cmd_value)

        elif action == "extract_excel_dom":
            result = extract_excel_value_dom(
                state.goal["excel"]["url"],
                state.goal["excel"]["label"]
            )
            print(">>> Excel DOM result:", result)
            log_step(state, action, result)
            if result.get("status") == "ok":
                state.excel_value = result["value"]

        elif action == "extract_excel_vision":
            result = extract_excel_value_vision(
                state.goal["excel"]["url"],
                state.goal["excel"]["label"]
            )
            print(">>> Excel vision result:", result)
            log_step(state, action, result)
            if result.get("status") == "ok":
                state.excel_value = result["value"]

        elif action == "compare":
            tol = state.goal["comparison_spec"]["tolerance"]
            state.comparison = {
                "match": abs(state.cmd_value - state.excel_value) <= tol,
                "cmd_value": state.cmd_value,
                "excel_value": state.excel_value,
                "tolerance": tol,
            }
            log_step(state, action, state.comparison)

        elif action == "finish":
            print("\n✅ FINAL RESULT")
            print(json.dumps(state.comparison, indent=2))
            log_step(state, action, "done")
            return

        elif action == "escalate":
            log_step(state, action, "escalated")
            human_escalation(state.__dict__, "Agent chose to escalate or could not complete reliably.")

        else:
            log_step(state, action, "unknown_action")
            human_escalation(state.__dict__, f"Unknown action: {action}")


# ---------------------------------------------------------------------

if __name__ == "__main__":
    user_prompt = r"""
Find TOTAL in:
"C:\Users\fatemeh.torabi.asr\small_vsc_projects\quick_python_project\quick_python_project\hybrid_orchestration_demo\data\invoice.txt"
and FV in the spreadsheet at
https://avanade-my.sharepoint.com/:x:/r/personal/fatemeh_torabi_asr_avanade_com/Documents/test_data_folder/Book.xlsx?d=w8f930cd74d9641d7a294db6ff4f350db&csf=1&web=1&e=klFwv8
Compare with tolerance 0.01.
"""
    run(user_prompt)
