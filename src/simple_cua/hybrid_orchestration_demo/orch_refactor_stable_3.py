"""
Vision-First Hybrid Orchestrator (Final, Decoupled)
==================================================

Agentic Roles:
- Orchestrator (run loop): executes actions and advances world state
- Orchestrator Agent (LLM): decides NEXT action only (task-agnostic)
- Task Agent: defines task state, parses goal, prepares task artifacts
- Tools: transactional, side-effecting skills (CMD, Excel, Vision)
"""

import os
import time
import json
import base64
import subprocess
import re
from dataclasses import dataclass, field
from typing import Optional, Dict, Any, List

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
    azure_endpoint=AZURE_ENDPOINT,
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


def human_escalation(world: dict, reason: str) -> None:
    print("\n⚠️ HUMAN ESCALATION REQUIRED")
    print("Reason:", reason)
    print("World state:")
    print(json.dumps(world, indent=2, default=str))
    raise RuntimeError(reason)


# ---------------------------------------------------------------------
# TASK AGENT — STATE (task-specific)
# ---------------------------------------------------------------------

@dataclass
class WorldState:
    goal: Optional[Dict[str, Any]] = None

    # CMD
    cmd_rendered: bool = False
    cmd_image: Optional[str] = None
    cmd_value: Optional[float] = None

    # Excel
    excel_value: Optional[float] = None
    excel_attempts: List[Dict[str, Any]] = field(default_factory=list)

    # Result
    comparison: Optional[Dict[str, Any]] = None

    # Episodic memory (generic)
    history: List[Dict[str, Any]] = field(default_factory=list)


# ---------------------------------------------------------------------
# TASK AGENT — GOAL PARSER
# ---------------------------------------------------------------------

def parse_goal(prompt: str) -> dict:
    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": (
                    "Extract task parameters.\n"
                    "Return ONLY JSON with EXACT schema:\n"
                    "{\n"
                    "  \"cmd\": { \"label\": string, \"file_path\": string },\n"
                    "  \"excel\": { \"label\": string, \"url\": string },\n"
                    "  \"comparison_spec\": { \"tolerance\": number }\n"
                    "}\n"
                    "Labels must be literal tokens only (e.g., TOTAL, FV)."
                ),
            },
            {"role": "user", "content": prompt},
        ],
    )

    parsed = json.loads(response.choices[0].message.content)
    parsed["cmd"]["label"] = parsed["cmd"]["label"].strip()
    parsed["excel"]["label"] = parsed["excel"]["label"].strip()
    return parsed


# ---------------------------------------------------------------------
# TASK AGENT — ARTIFACT PREP (MINIMAL ADDITION)
# ---------------------------------------------------------------------

def normalize_excel_url(url: str) -> str:
    """
    Ensure Excel opens in web viewer mode.
    """
    if "web=1" not in url:
        sep = "&" if "?" in url else "?"
        return f"{url}{sep}web=1"
    return url


# ---------------------------------------------------------------------
# TOOLS — CMD
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
# TOOL — VISION
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
                    {"type": "text", "text": f"Extract numeric value near label '{label}'."},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_b64}"}},
                ],
            },
        ],
    )

    parsed = json.loads(response.choices[0].message.content)
    print(f">>> Vision extracted value for '{label}': {parsed['value']}")
    return float(parsed["value"])


# ---------------------------------------------------------------------
# TOOLS — EXCEL (UNCHANGED)
# ---------------------------------------------------------------------


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
# ORCHESTRATOR AGENT (GENERIC LLM BRAIN)
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


def decide_next_action(world: WorldState) -> str:
    print("\n>>> [Orchestrator Agent] deciding next action...")
    messages=[
            {
                "role": "system",
                "content": (
                    "You are an Orchestrator Agent.\n"
                    "Decide the NEXT action only.\n"
                    "Do not repeat actions that would not change the world.\n"
                    "If stuck or uncertain, choose escalate.\n"
                    "Return JSON: { \"action\": string }"
                ),
            },
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "goal": world.goal,
                        "world_state": world.__dict__,
                        "recent_steps": world.history[-4:],
                        "available_actions": ACTIONS,
                    },
                    indent=2,
                    default=str,
                ),
            },
        ]
    
    print(">>> [Orchestrator Agent] messages:\n", messages)
    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        response_format={"type": "json_object"},
        messages=messages,
    )
    
    print(">>> [Orchestrator Agent] raw response:", response.choices[0].message.content)

    decision = safe_json_parse(response.choices[0].message.content)
    return decision.get("action", "escalate")


# ---------------------------------------------------------------------
# ORCHESTRATOR LOOP
# ---------------------------------------------------------------------

def run(prompt: str):
    world = WorldState(goal=parse_goal(prompt))

    # ✅ APPLY URL NORMALIZATION ONCE (TASK LAYER)
    world.goal["excel"]["url"] = normalize_excel_url(world.goal["excel"]["url"])

    print("\n>>> Parsed goal:\n", json.dumps(world.goal, indent=2))

    while True:
        action = decide_next_action(world)
        print(f"\n🧠 Orchestrator decided: {action}")

        # ✅ Deduplicated episodic memory
        if not world.history or world.history[-1]["action"] != action:
            world.history.append({"action": action})

        if action == "render_cmd_file":
            render_file_in_cmd(world.goal["cmd"]["file_path"])
            world.cmd_rendered = True

        elif action == "capture_cmd":
            world.cmd_image = capture_cmd_screenshot()

        elif action == "extract_cmd_value":
            world.cmd_value = extract_numeric_value_near_label(
                world.cmd_image, world.goal["cmd"]["label"]
            )

        elif action == "extract_excel_dom":
            res = extract_excel_value_dom(
                world.goal["excel"]["url"], world.goal["excel"]["label"]
            )
            world.excel_attempts.append({"mode": "dom", "result": res})
            if res["status"] == "ok":
                world.excel_value = res["value"]

        elif action == "extract_excel_vision":
            res = extract_excel_value_vision(
                world.goal["excel"]["url"], world.goal["excel"]["label"]
            )
            world.excel_attempts.append({"mode": "vision", "result": res})
            if res["status"] == "ok":
                world.excel_value = res["value"]

        elif action == "compare":
            tol = world.goal["comparison_spec"]["tolerance"]
            world.comparison = {
                "match": abs(world.cmd_value - world.excel_value) <= tol,
                "cmd_value": world.cmd_value,
                "excel_value": world.excel_value,
            }

        elif action == "finish":
            print("\n✅ FINAL RESULT")
            print(json.dumps(world.comparison, indent=2))
            return

        elif action == "escalate":
            human_escalation(world.__dict__, "Orchestrator escalated")

        else:
            human_escalation(world.__dict__, f"Unknown action: {action}")


# ---------------------------------------------------------------------

if __name__ == "__main__":
    user_prompt = r"""
Find TOTAL in:
"C:\Users\fatemeh.torabi.asr\small_vsc_projects\quick_python_project\quick_python_project\hybrid_orchestration_demo\data\invoice.txt"
and FV in the spreadsheet at
https://avanade-my.sharepoint.com/:x:/r/personal/fatemeh_torabi_asr_avanade_com/Documents/test_data_folder/Book.xlsx
Compare with tolerance 0.01.
"""
    run(user_prompt)
