"""
Vision-First Hybrid Orchestrator (Refactored)
============================================

Goal:
- Extract numeric value near a label from a local text file rendered in CMD (vision-only)
- Extract numeric value near a label from an Excel Online workbook using TWO independent strategies:
    1) DOM-based (best effort; often fails for Excel Online grids)
    2) Vision-based (screenshot + vision extraction)
- Compare the two values within tolerance
- Brain (LLM) chooses the next action at runtime (DOM vs Vision, retries, etc.)
- Tools are transactional (no persistent Excel handler context across tool calls)
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
    excel_attempts: list = field(default_factory=list)

    # Final
    comparison: Optional[Dict[str, Any]] = None
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
    # Render settle time (tune as needed)
    page.wait_for_timeout(12_000)

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

# Partially deterministic brain decision function
# def decide_next_action(state: State) -> str:
#     """
#     LLM-based brain decision with hard guardrails to prevent loop traps.
#     The LLM chooses between DOM and vision for Excel, but cannot:
#       - redo CMD steps once cmd_value exists
#       - compare/finish prematurely
#       - retry the same Excel mode infinitely
#     """
#     # Derived facts
#     dom_tried = any(a.get("mode") == "dom" for a in state.excel_attempts)
#     vision_tried = any(a.get("mode") == "vision" for a in state.excel_attempts)

#     # Hard guards (monotonic progress)
#     if not state.cmd_rendered:
#         return "render_cmd_file"
#     if state.cmd_image is None:
#         return "capture_cmd"
#     if state.cmd_value is None:
#         return "extract_cmd_value"

#     if state.excel_value is not None and state.comparison is None:
#         return "compare"
#     if state.comparison is not None:
#         return "finish"

#     # If Excel not extracted yet: let LLM choose DOM vs vision, but bounded.
#     if state.excel_value is None:
#         # If both tried, escalate.
#         if dom_tried and vision_tried:
#             return "escalate"

#         # Let the LLM choose next Excel strategy.
#         state_summary = {
#             "cmd_value_present": state.cmd_value is not None,
#             "excel_value_present": state.excel_value is not None,
#             "dom_tried": dom_tried,
#             "vision_tried": vision_tried,
#             "excel_attempts": state.excel_attempts[-2:],  # last two
#             "allowed_actions": ["extract_excel_dom", "extract_excel_vision", "escalate"],
#         }

#         response = client.chat.completions.create(
#             model=DEPLOYMENT_NAME,
#             response_format={"type": "json_object"},
#             messages=[
#                 {
#                     "role": "system",
#                     "content": (
#                         "You are the Orchestrator Brain.\n"
#                         "Choose the NEXT action for EXCEL extraction.\n\n"
#                         "You may choose:\n"
#                         "- extract_excel_dom\n"
#                         "- extract_excel_vision\n"
#                         "- escalate\n\n"
#                         "Rules:\n"
#                         "- If dom_tried is true, avoid choosing extract_excel_dom again.\n"
#                         "- If vision_tried is true, avoid choosing extract_excel_vision again.\n"
#                         "- Prefer trying the untried method first.\n"
#                         "- Escalate if you believe both are unlikely to work.\n\n"
#                         "Return JSON: { \"action\": string }"
#                     ),
#                 },
#                 {"role": "user", "content": json.dumps(state_summary, indent=2)},
#             ],
#         )

#         decision = json.loads(response.choices[0].message.content)
#         action = decision.get("action")

#         if action not in {"extract_excel_dom", "extract_excel_vision", "escalate"}:
#             return "escalate"

#         # Enforce non-retry of the same method
#         if action == "extract_excel_dom" and dom_tried:
#             return "extract_excel_vision" if not vision_tried else "escalate"
#         if action == "extract_excel_vision" and vision_tried:
#             return "extract_excel_dom" if not dom_tried else "escalate"

#         return action

#     return "escalate"


# # llm-based brain decision function with mechanical gaurdrails
# def decide_next_action(state: State) -> str:
#     """
#     LLM-based orchestrator brain.

#     - Brain can choose among ALL actions each step.
#     - We pass explicit monotonic rules so it doesn't repeat actions.
#     - We still validate the action locally and override if it's invalid
#       (prevents loop traps and bad decisions).
#     """

#     # ----- Build a compact state summary (LLM-friendly) -----
#     state_summary = {
#         # CMD
#         "cmd_rendered": state.cmd_rendered,
#         "cmd_image_present": state.cmd_image is not None,
#         "cmd_value_present": state.cmd_value is not None,

#         # Excel
#         "excel_value_present": state.excel_value is not None,
#         "excel_dom_tried": any(a.get("mode") == "dom" for a in state.excel_attempts),
#         "excel_vision_tried": any(a.get("mode") == "vision" for a in state.excel_attempts),
#         "last_excel_attempt": state.excel_attempts[-1] if state.excel_attempts else None,

#         # Final
#         "comparison_present": state.comparison is not None,
#         "errors_count": len(state.errors),
#     }

#     allowed_actions = ALLOWED_ACTIONS  # uses your global list

#     # ----- LLM decides -----
#     response = client.chat.completions.create(
#         model=DEPLOYMENT_NAME,
#         response_format={"type": "json_object"},
#         messages=[
#             {
#                 "role": "system",
#                 "content": (
#                     "You are the Orchestrator Brain for a vision-first automation.\n"
#                     "You decide the NEXT action only.\n\n"

#                     "Available actions:\n"
#                     "- render_cmd_file\n"
#                     "- capture_cmd\n"
#                     "- extract_cmd_value\n"
#                     "- extract_excel_dom\n"
#                     "- extract_excel_vision\n"
#                     "- compare\n"
#                     "- finish\n"
#                     "- escalate\n\n"

#                     "STRICT RULES (do not violate):\n\n"

#                     "CMD progression:\n"
#                     "1) render_cmd_file ONLY if cmd_rendered is false.\n"
#                     "2) capture_cmd ONLY if cmd_rendered is true AND cmd_image_present is false.\n"
#                     "3) extract_cmd_value ONLY if cmd_image_present is true AND cmd_value_present is false.\n"
#                     "4) Never go back to earlier CMD actions once cmd_value_present is true.\n\n"

#                     "Excel extraction:\n"
#                     "- extract_excel_dom ONLY if excel_value_present is false.\n"
#                     "- extract_excel_vision ONLY if excel_value_present is false.\n"
#                     "- Prefer NOT to repeat the same excel method if it already failed.\n"
#                     "- If one method failed, try the other next.\n"
#                     "- Escalate if both methods were tried and failed.\n\n"

#                     "Finalization:\n"
#                     "- compare ONLY if cmd_value_present is true AND excel_value_present is true AND comparison_present is false.\n"
#                     "- finish ONLY if comparison_present is true.\n\n"

#                     "General:\n"
#                     "- Choose exactly ONE next action.\n"
#                     "- Do NOT repeat actions that would not change state.\n"
#                     "- If unsure or stuck, choose escalate.\n\n"

#                     "Return ONLY JSON: {\"action\": string, \"reason\": string}"
#                 ),
#             },
#             {
#                 "role": "user",
#                 "content": json.dumps(
#                     {
#                         "goal": state.goal,
#                         "state_summary": state_summary,
#                         "allowed_actions": allowed_actions,
#                         "recent_excel_attempts": state.excel_attempts[-2:],  # keep it short
#                     },
#                     indent=2,
#                     default=str
#                 ),
#             },
#         ],
#     )

#     decision = json.loads(response.choices[0].message.content)
#     action = decision.get("action")
#     reason = decision.get("reason", "")

#     print("\n🧠 Brain rationale:", reason)

#     # ----- Local validation / guardrail enforcement -----
#     def invalid(a: str) -> bool:
#         return a not in allowed_actions

#     # Hard monotonic guards:
#     if invalid(action):
#         return "escalate"

#     if action == "render_cmd_file" and state.cmd_rendered:
#         return "escalate"  # violates monotonicity

#     if action == "capture_cmd" and (not state.cmd_rendered or state.cmd_image is not None):
#         return "escalate"

#     if action == "extract_cmd_value" and (state.cmd_image is None or state.cmd_value is not None):
#         return "escalate"

#     # Once cmd_value exists, never redo cmd steps
#     if state.cmd_value is not None and action in {"render_cmd_file", "capture_cmd", "extract_cmd_value"}:
#         return "escalate"

#     # Excel actions only if excel_value missing
#     if action in {"extract_excel_dom", "extract_excel_vision"} and state.excel_value is not None:
#         return "escalate"

#     # Compare/finish rules
#     if action == "compare":
#         if state.cmd_value is None or state.excel_value is None or state.comparison is not None:
#             return "escalate"

#     if action == "finish" and state.comparison is None:
#         return "escalate"

#     # Escalate if both excel paths already failed and LLM tries again
#     dom_tried = any(a.get("mode") == "dom" for a in state.excel_attempts)
#     vision_tried = any(a.get("mode") == "vision" for a in state.excel_attempts)
#     if state.excel_value is None and dom_tried and vision_tried and action in {"extract_excel_dom", "extract_excel_vision"}:
#         return "escalate"

#     return action


# Maximally llm-based brain decision function 
def decide_next_action(state: State) -> str:
    """
    Pure LLM brain:
    - All workflow rules are enforced by the prompt.
    - No external state-based guards (only minimal parsing fallback).
    """

    state_summary = {
        # CMD
        "cmd_rendered": state.cmd_rendered,
        "cmd_image_present": state.cmd_image is not None,
        "cmd_value_present": state.cmd_value is not None,

        # Excel
        "excel_value_present": state.excel_value is not None,
        "excel_dom_tried": any(a.get("mode") == "dom" for a in state.excel_attempts),
        "excel_vision_tried": any(a.get("mode") == "vision" for a in state.excel_attempts),
        "last_excel_attempt": state.excel_attempts[-1] if state.excel_attempts else None,

        # Final
        "comparison_present": state.comparison is not None,
        "errors_count": len(state.errors),
    }

    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": (
                    "You are the Orchestrator Brain for a vision-first automation.\n"
                    "You are tasked with completing a multi-step workflow involving CMD rendering, Excel extraction, and final comparison.\n"
                    "The goal is to extract numeric values from CMD and Excel, then compare them within a specified tolerance and report the result.\n"
                    "At each step, you are asked to decide the NEXT action only given the current state.\n\n"

                    "Available actions:\n"
                    "- render_cmd_file\n"
                    "- capture_cmd\n"
                    "- extract_cmd_value\n"
                    "- extract_excel_dom\n"
                    "- extract_excel_vision\n"
                    "- compare\n"
                    "- finish\n"
                    "- escalate\n\n"

                    "RULES (follow strictly):\n"
                    "CMD progression:\n"
                    "1) render_cmd_file ONLY if cmd_rendered is false.\n"
                    "2) capture_cmd ONLY if cmd_rendered is true AND cmd_image_present is false.\n"
                    "3) extract_cmd_value ONLY if cmd_image_present is true AND cmd_value_present is false.\n"
                    "4) Once cmd_value_present is true, DO NOT choose any CMD actions again.\n\n"

                    "Excel extraction:\n"
                    "- extract_excel_dom ONLY if excel_value_present is false.\n"
                    "- extract_excel_vision ONLY if excel_value_present is false.\n"
                    "- If one Excel method fails or returns not_found, try the other method next.\n"
                    "- Do not repeat the same Excel method if it already failed.\n"
                    "- If both Excel methods were tried and excel_value is still missing, choose escalate.\n\n"

                    "Finalization:\n"
                    "- compare ONLY if cmd_value_present AND excel_value_present AND comparison_present is false.\n"
                    "- finish ONLY if comparison_present is true.\n\n"

                    "General:\n"
                    "- Never repeat an action that would not change the state.\n"
                    "- If uncertain or stuck, choose escalate.\n\n"

                    "Return JSON only:\n"
                    "{ \"action\": string, \"reason\": string }"
                ),
            },
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "goal": state.goal,
                        "state_summary": state_summary,
                        "allowed_actions": ALLOWED_ACTIONS,
                        "recent_excel_attempts": state.excel_attempts[-3:],
                    },
                    indent=2,
                    default=str
                ),
            },
        ],
    )

    # Minimal fallback only (no extra guards)
    try:
        decision = json.loads(response.choices[0].message.content)
        action = decision.get("action")
        reason = decision.get("reason", "")
        print("\n🧠 Brain rationale:", reason)

        if action not in ALLOWED_ACTIONS:
            print(">>> Brain returned invalid action; escalating.")
            return "escalate"

        return action
    except Exception as e:
        print(">>> Could not parse brain output; escalating:", repr(e))
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

        elif action == "capture_cmd":
            state.cmd_image = capture_cmd_screenshot()

        elif action == "extract_cmd_value":
            state.cmd_value = extract_numeric_value_near_label(
                state.cmd_image,
                state.goal["cmd"]["label"]
            )

        elif action == "extract_excel_dom":
            result = extract_excel_value_dom(
                state.goal["excel"]["url"],
                state.goal["excel"]["label"]
            )
            print(">>> Excel DOM result:", result)
            state.excel_attempts.append({"mode": "dom", "result": result})
            if result["status"] == "ok":
                state.excel_value = result["value"]

        elif action == "extract_excel_vision":
            result = extract_excel_value_vision(
                state.goal["excel"]["url"],
                state.goal["excel"]["label"]
            )
            print(">>> Excel vision result:", result)
            state.excel_attempts.append({"mode": "vision", "result": result})
            if result["status"] == "ok":
                state.excel_value = result["value"]

        elif action == "compare":
            tol = state.goal["comparison_spec"]["tolerance"]
            state.comparison = {
                "match": abs(state.cmd_value - state.excel_value) <= tol,
                "cmd_value": state.cmd_value,
                "excel_value": state.excel_value,
                "tolerance": tol,
                "excel_attempts": state.excel_attempts,
            }

        elif action == "finish":
            print("\n✅ FINAL RESULT")
            print(json.dumps(state.comparison, indent=2))
            return

        elif action == "escalate":
            human_escalation(state.__dict__, "Failed to extract Excel value via DOM and/or Vision.")

        else:
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
