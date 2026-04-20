"""
Vision-First Hybrid Orchestrator
===============================

Supports multiple tasks using:
- A task-agnostic Orchestrator (LLM)
- Task-specific Task Agents
- Shared CMD / Excel / Vision tools

Tasks implemented:
1) CompareTaskAgent: extract CMD value + Excel value → compare
2) CopyPasteTaskAgent: extract CMD value → write to Excel cell
"""

import os
import time
import json
import base64
import subprocess
import re
from dataclasses import dataclass, field
from typing import Dict, Any, Optional, List

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
# UTILS
# ---------------------------------------------------------------------

def safe_json_parse(text: str) -> dict:
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*", "", text)
        text = text.replace("```", "").strip()
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise ValueError(f"No JSON found:\n{text}")
    return json.loads(match.group(0))


def human_escalation(world: dict, reason: str):
    print("\n⚠️ HUMAN ESCALATION")
    print("Reason:", reason)
    print(json.dumps(world, indent=2, default=str))
    raise RuntimeError(reason)


def log_world(world: "WorldState", step: int, action: str, rationale: str = ""):
    """Debug-friendly compact snapshot each step."""
    print("\n" + "─" * 70)
    print(f"STEP {step} | action={action}")
    if rationale:
        print(f"rationale: {rationale}")
    print(f"values:   {sorted(list(world.values.keys()))}")
    print(f"artifacts:{sorted(list(world.artifacts.keys()))}")
    if world.errors:
        print(f"errors:   {world.errors[-3:]}")
    if world.history:
        print(f"history(last 6): {[h.get('action') for h in world.history[-6:]]}")
    print("─" * 70 + "\n")


# ---------------------------------------------------------------------
# WORLD STATE (TASK-AGNOSTIC)
# ---------------------------------------------------------------------

@dataclass
class WorldState:
    goal: dict
    values: Dict[str, Any] = field(default_factory=dict)
    artifacts: Dict[str, Any] = field(default_factory=dict)
    history: List[dict] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    done: bool = False

# ---------------------------------------------------------------------
# TOOLS (UNCHANGED / SHARED)
# ---------------------------------------------------------------------

# ---------- CMD ----------

def render_file_in_cmd(file_path: str):
    """
    TOOL: CMD render
    Opens a new Command Prompt window (detached),
    waits for it to appear, then types the file contents.
    """
    subprocess.Popen("start cmd.exe", shell=True)
    time.sleep(1)

    win = None
    for _ in range(20):  # wait up to ~10 seconds
        wins = [
            w for w in gw.getAllWindows()
            if "Command Prompt" in w.title or "cmd" in w.title.lower()
        ]
        if wins:
            win = wins[0]
            break
        time.sleep(0.5)

    if win is None:
        raise RuntimeError("CMD window did not appear")

    win.activate()
    time.sleep(0.5)

    pyautogui.write(f'type "{file_path}"', interval=0.03)
    pyautogui.press("enter")
    time.sleep(1.5)


def capture_cmd_screenshot() -> str:
    """
    TOOL: CMD capture
    Waits for a visible Command Prompt window and screenshots it.
    """
    win = None
    for _ in range(20):  # wait up to ~10 seconds
        wins = [
            w for w in gw.getAllWindows()
            if "Command Prompt" in w.title or "cmd" in w.title.lower()
        ]
        if wins:
            win = wins[0]
            break
        time.sleep(0.5)

    if win is None:
        raise RuntimeError("CMD window not found for screenshot")

    win.activate()
    time.sleep(0.3)

    bbox = (win.left, win.top, win.right, win.bottom)
    img = ImageGrab.grab(bbox=bbox)
    img.save(CMD_SCREENSHOT_PATH)

    print(f">>> CMD screenshot saved to {CMD_SCREENSHOT_PATH}")
    return CMD_SCREENSHOT_PATH


def extract_numeric_value_near_label(image_path: str, label: str) -> float:
    with open(image_path, "rb") as f:
        img_b64 = base64.b64encode(f.read()).decode()

    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": "Return JSON: {\"value\": number}"},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": f"Extract numeric value near label '{label}'."},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_b64}"}},
                ],
            },
        ],
    )
    val = float(json.loads(response.choices[0].message.content)["value"])
    print(f">>> Vision extracted '{label}' => {val}")
    return val

# ---------- EXCEL (same stable code you validated) ----------

def _open_excel_page(p, excel_url: str):
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
    return context, page


def extract_excel_value_dom(excel_url: str, label: str) -> dict:
    try:
        with sync_playwright() as p:
            ctx, page = _open_excel_page(p, excel_url)
            page.screenshot(path=EXCEL_SCREENSHOT_PATH, full_page=True)

            probe = page.evaluate(
                """
                (label) => {
                  const elems = document.querySelectorAll('div, span, td, [role="gridcell"]');
                  for (const e of elems) {
                    const t = (e.innerText || '').trim();
                    if (t === label) {
                      const sib = e.nextElementSibling;
                      return sib ? (sib.innerText || '').trim() : null;
                    }
                  }
                  return null;
                }
                """,
                label,
            )
            ctx.close()

            if probe is None:
                return {"status": "not_found"}
            return {"status": "ok", "value": float(probe)}

    except Exception as e:
        return {"status": "error", "reason": str(e)}


def extract_excel_value_vision(excel_url: str, label: str) -> dict:
    try:
        with sync_playwright() as p:
            ctx, page = _open_excel_page(p, excel_url)
            page.screenshot(path=EXCEL_SCREENSHOT_PATH, full_page=True)
            ctx.close()

        val = extract_numeric_value_near_label(EXCEL_SCREENSHOT_PATH, label)
        return {"status": "ok", "value": val}

    except Exception as e:
        return {"status": "error", "reason": str(e)}

# ---------------------------------------------------------------------
# TASK AGENTS (TASK-SPECIFIC)
# ---------------------------------------------------------------------

# -------------------------------
# Task Agent: Compare Values
# -------------------------------
def compare_task_agent(world: WorldState):
    """
    ROLE: Task Agent
    Consumes: world.values["cmd_value"], world.values["excel_value"]
    Produces: world.values["comparison"], sets world.done=True
    """
    tol = world.goal["comparison_spec"]["tolerance"]

    if "cmd_value" not in world.values or "excel_value" not in world.values:
        world.errors.append("compare_task_agent called without required values")
        return

    cmd_v = world.values["cmd_value"]
    excel_v = world.values["excel_value"]

    world.values["comparison"] = {
        "match": abs(cmd_v - excel_v) <= tol,
        "cmd_value": cmd_v,
        "excel_value": excel_v,
        "tolerance": tol,
    }

    world.done = True


def copy_paste_task_agent(world: WorldState):
    """
    ROLE: Task Agent (Copy CMD → Excel)
    (Excel write is stubbed here)
    """
    if "cmd_value" not in world.values:
        world.errors.append("copy_paste_task_agent called without cmd_value")
        return

    # Stub: simulate Excel write
    world.values["excel_written"] = True
    world.done = True

# ---------------------------------------------------------------------
# ORCHESTRATOR (TASK-AGNOSTIC)
# ---------------------------------------------------------------------

ACTIONS = [
    "render_cmd_file",
    "capture_cmd",
    "extract_cmd_value",
    "extract_excel_dom",
    "extract_excel_vision",
    "compare",
    "write_excel",
    "finish",
    "escalate",
]


def decide_next_action(world: WorldState) -> dict:
    """
    ROLE: Orchestrator (task-agnostic)
    Returns dict: {"action": str, "rationale": str}
    """
    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": (
                    "You are an Orchestrator for a multi-step automation.\n"
                    "You do not know the task domain.\n"
                    "Your job is to select the NEXT action that advances progress.\n\n"
                    "Rules:\n"
                    "- Never repeat an action that already succeeded.\n"
                    "- If no progress seems possible, escalate.\n"
                    "- Prefer actions that create missing prerequisites before dependent actions.\n\n"
                    "Return JSON: {\"action\": string, \"rationale\": string}"
                ),
            },
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "known_values": sorted(list(world.values.keys())),
                        "known_artifacts": sorted(list(world.artifacts.keys())),
                        "recent_history": world.history[-6:],
                        "allowed_actions": ACTIONS,
                    },
                    indent=2,
                ),
            },
        ],
    )

    try:
        obj = json.loads(response.choices[0].message.content)
        if obj.get("action") not in ACTIONS:
            return {"action": "escalate", "rationale": "Invalid action from model"}
        return {"action": obj.get("action"), "rationale": obj.get("rationale", "")}
    except Exception as e:
        return {"action": "escalate", "rationale": f"Could not parse model output: {repr(e)}"}

# ---------------------------------------------------------------------
# GOAL PARSER
# ---------------------------------------------------------------------

def parse_goal(prompt: str) -> dict:
    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": (
                    "Extract structured task parameters.\n"
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
    return json.loads(response.choices[0].message.content)

# ---------------------------------------------------------------------
# RUN LOOP
# ---------------------------------------------------------------------

def run(task_agent):
    goal = parse_goal(USER_PROMPT)
    world = WorldState(goal=goal)

    step = 0
    while True:
        step += 1

        if world.done:
            print("\n✅ FINAL WORLD STATE (values)")
            print(json.dumps(world.values, indent=2, default=str))
            print("\n✅ FINAL WORLD STATE (artifacts)")
            print(json.dumps(world.artifacts, indent=2, default=str))
            return

        decision = decide_next_action(world)
        action = decision["action"]
        rationale = decision.get("rationale", "")

        world.history.append({"action": action, "rationale": rationale})

        log_world(world, step, action, rationale)

        try:
            if action == "render_cmd_file":
                render_file_in_cmd(goal["cmd"]["file_path"])
                world.artifacts["cmd_rendered"] = True

            elif action == "capture_cmd":
                world.artifacts["cmd_image"] = capture_cmd_screenshot()

            elif action == "extract_cmd_value":
                if "cmd_image" not in world.artifacts:
                    world.errors.append("extract_cmd_value attempted without cmd_image")
                    continue
                world.values["cmd_value"] = extract_numeric_value_near_label(
                    world.artifacts["cmd_image"], goal["cmd"]["label"]
                )

            elif action == "extract_excel_dom":
                res = extract_excel_value_dom(goal["excel"]["url"], goal["excel"]["label"])
                world.artifacts.setdefault("excel_attempts", []).append({"mode": "dom", "result": res})
                if res.get("status") == "ok":
                    world.values["excel_value"] = res["value"]

            elif action == "extract_excel_vision":
                res = extract_excel_value_vision(goal["excel"]["url"], goal["excel"]["label"])
                world.artifacts.setdefault("excel_attempts", []).append({"mode": "vision", "result": res})
                if res.get("status") == "ok":
                    world.values["excel_value"] = res["value"]

            elif action == "compare":
                task_agent(world)

            elif action == "write_excel":
                task_agent(world)

            elif action == "finish":
                print("\n✅ FINAL WORLD STATE (values)")
                print(json.dumps(world.values, indent=2, default=str))
                return

            elif action == "escalate":
                human_escalation(world.__dict__, "Orchestrator escalated")

        except Exception as e:
            world.errors.append(f"{action} failed: {repr(e)}")
            # let orchestrator decide next action (retry/alternate/escalate)

# ---------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------

USER_PROMPT = r"""
Find TOTAL in:
"C:\Users\fatemeh.torabi.asr\small_vsc_projects\quick_python_project\quick_python_project\hybrid_orchestration_demo\data\invoice.txt"
and FV in:
https://avanade-my.sharepoint.com/:x:/r/personal/fatemeh_torabi_asr_avanade_com/Documents/test_data_folder/Book.xlsx
Compare with tolerance 0.01.
"""

if __name__ == "__main__":
    # 🔁 SWITCH TASK HERE
    run(compare_task_agent)
    # run(copy_paste_task_agent)
