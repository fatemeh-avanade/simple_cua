"""
Hybrid LLM-Brained Orchestrator
==============================

Example user prompt:

"Look up the value labeled TOTAL in the file at
 C:\\Users\\me\\data\\invoice.txt using the command line.
 Look up the value labeled FV in the Excel sheet at
 https://<sharepoint-url>/Book.xlsx.
 Compare the two values and tell me if they match within tolerance 0.01."

Architecture:
- GPT-4o = orchestrator brain (planning + decisions)
- TerminalAgent = pyautogui + screenshot
- VisionAgent = GPT-4o image understanding
- ExcelAgent = Playwright (browser automation)
- Human-in-the-loop for escalation
"""

import os
import time
import json
import base64
import subprocess
import pyautogui
import pygetwindow as gw
from PIL import ImageGrab
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv
from openai import AzureOpenAI

# ---------------------------------------------------------------------
# ENV & CONFIG
# ---------------------------------------------------------------------

load_dotenv()

DATA_PATH = r"C:\Users\fatemeh.torabi.asr\small_vsc_projects\quick_python_project\quick_python_project\hybrid_orchestration_demo\data"
SCREENSHOT_PATH = os.path.join(DATA_PATH, "terminal_screen.png")

EDGE_PROFILE = r"C:\Users\fatemeh.torabi.asr\AppData\Local\Microsoft\Edge\User Data\Work"

AZURE_ENDPOINT = "https://fa-test-openai-instance-canada-east.openai.azure.com/"
DEPLOYMENT_NAME = "fa-test-gpt-4o"

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
    text = text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
    return json.loads(text)

# ---------------------------------------------------------------------
# AGENTS
# ---------------------------------------------------------------------

class TerminalAgent:
    def open_and_capture(self, file_path: str) -> dict:
        try:
            subprocess.Popen("start cmd.exe", shell=True)
            time.sleep(2)

            pyautogui.write(f"type {file_path}\n", interval=0.05)
            time.sleep(1)

            wins = gw.getWindowsWithTitle("cmd.exe")
            if not wins:
                raise RuntimeError("CMD window not found")

            win = wins[0]
            bbox = (win.left, win.top, win.right, win.bottom)
            screenshot = ImageGrab.grab(bbox=bbox)
            screenshot.save(SCREENSHOT_PATH)

            return {"status": "ok", "screenshot": SCREENSHOT_PATH}

        except Exception as e:
            return {"status": "error", "error": str(e)}

class VisionAgent:
    def extract_labeled_value(self, image_path: str, label: str) -> dict:
        try:
            with open(image_path, "rb") as f:
                img_b64 = base64.b64encode(f.read()).decode("utf-8")

            response = client.chat.completions.create(
                model=DEPLOYMENT_NAME,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": (
                                f"Read the screenshot and extract the numeric value "
                                f"associated with the label '{label}'. "
                                f"Return ONLY valid JSON like {{\"value\": 123.45}}"
                            )
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{img_b64}"
                            }
                        }
                    ]
                }]
            )

            parsed = safe_json_parse(response.choices[0].message.content)
            return {"status": "ok", "value": float(parsed["value"])}

        except Exception as e:
            return {"status": "error", "error": str(e)}

class ExcelAgent:
    def read_labeled_value(self, label: str, excel_url: str) -> dict:
        try:
            with sync_playwright() as p:
                ctx = p.chromium.launch_persistent_context(
                    EDGE_PROFILE,
                    channel="msedge",
                    headless=False,
                    args=["--no-first-run", "--disable-extensions"]
                )
                page = ctx.new_page()
                page.goto(excel_url)
                page.wait_for_timeout(5000)

                value = page.evaluate(f"""
                    () => {{
                        const elems = document.querySelectorAll('*');
                        for (const e of elems) {{
                            const t = (e.innerText || '').trim();
                            if (t === '{label}') {{
                                const sib = e.nextElementSibling;
                                if (sib) return sib.innerText.trim();
                            }}
                        }}
                        return null;
                    }}
                """)

                ctx.close()

                if value is None:
                    raise RuntimeError(f"Label '{label}' not found in Excel")

                return {"status": "ok", "value": float(value)}

        except Exception as e:
            return {"status": "error", "error": str(e)}

# ---------------------------------------------------------------------
# HUMAN-IN-THE-LOOP
# ---------------------------------------------------------------------

def human_escalation(context: dict) -> bool:
    print("\n⚠️ HUMAN ESCALATION REQUIRED")
    print(json.dumps(context, indent=2))
    return input("Approve override? (y/n): ").strip().lower() == "y"

# ---------------------------------------------------------------------
# ORCHESTRATOR (GPT-4o BRAIN)
# ---------------------------------------------------------------------

ALLOWED_ACTIONS = [
    "open_terminal_and_capture",
    "extract_labeled_value_from_image",
    "read_labeled_value_from_excel",
    "compare_values",
    "escalate_to_human",
    "finish"
]

def parse_user_goal(user_prompt: str) -> dict:
    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        messages=[
            {
                "role": "system",
                "content": (
                    "Extract structured task parameters from the user request.\n"
                    "Return ONLY valid JSON with keys:\n"
                    "terminal.file_path, terminal.label,\n"
                    "excel.url, excel.label,\n"
                    "comparison.tolerance"
                )
            },
            {"role": "user", "content": user_prompt}
        ]
    )
    return safe_json_parse(response.choices[0].message.content)

def decide_next_action(state: dict) -> dict:
    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        messages=[
            {
                "role": "system",
                "content": (
                    "You are the Orchestrator Agent.\n"
                    "Decide the NEXT action only.\n"
                    "You never execute UI actions yourself.\n"
                    "Return ONLY valid JSON with fields:\n"
                    "{ action, reason }"
                )
            },
            {
                "role": "user",
                "content": json.dumps({
                    "goal": state["goal"],
                    "state": state,
                    "allowed_actions": ALLOWED_ACTIONS
                })
            }
        ]
    )
    return safe_json_parse(response.choices[0].message.content)

# ---------------------------------------------------------------------
# MAIN ORCHESTRATION LOOP
# ---------------------------------------------------------------------

def run_orchestrator(user_prompt: str):
    terminal = TerminalAgent()
    vision = VisionAgent()
    excel = ExcelAgent()

    state = {
        "goal": parse_user_goal(user_prompt),
        "terminal": None,
        "vision": None,
        "excel": None,
        "comparison": None,
        "errors": []
    }

    while True:
        decision = decide_next_action(state)
        print("\n🧠 Orchestrator decision:")
        print(json.dumps(decision, indent=2))

        action = decision["action"]

        if action == "open_terminal_and_capture":
            state["terminal"] = terminal.open_and_capture(
                state["goal"]["terminal"]["file_path"]
            )

        elif action == "extract_labeled_value_from_image":
            state["vision"] = vision.extract_labeled_value(
                state["terminal"]["screenshot"],
                state["goal"]["terminal"]["label"]
            )

        elif action == "read_labeled_value_from_excel":
            state["excel"] = excel.read_labeled_value(
                state["goal"]["excel"]["label"],
                state["goal"]["excel"]["url"]
            )

        elif action == "compare_values":
            v1 = state["vision"]["value"]
            v2 = state["excel"]["value"]
            tol = state["goal"]["comparison"]["tolerance"]

            state["comparison"] = {
                "match": abs(v1 - v2) <= tol,
                "value_1": v1,
                "value_2": v2,
                "tolerance": tol
            }

        elif action == "escalate_to_human":
            if not human_escalation(state):
                print("❌ Task aborted.")
                return
            print("🧑‍⚖️ Human override approved.")
            return

        elif action == "finish":
            print("\n✅ FINAL RESULT")
            print(json.dumps(state["comparison"], indent=2))
            return

        else:
            print("❌ Unknown action. Escalating.")
            human_escalation(state)
            return

# ---------------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------------

if __name__ == "__main__":
    print("=== Hybrid LLM-Brained Orchestrator Demo ===")
    # user_prompt = input("\nEnter your task:\n> ")
    user_prompt = r"""
First look up the value labeled TOTAL in the file at

"C:\Users\fatemeh.torabi.asr\small_vsc_projects\quick_python_project\quick_python_project\hybrid_orchestration_demo\data\invoice.txt"

using the command line (open using the cmd tool and the tool for labeled value extraction from screenshot).

Then look up the value labeled FV in the Excel sheet at

https://avanade-my.sharepoint.com/:x:/r/personal/fatemeh_torabi_asr_avanade_com/Documents/test_data_folder/Book.xlsx?d=w8f930cd74d9641d7a294db6ff4f350db&csf=1&web=1&e=klFwv8

using your Excel reading tool.

Finally, compare the two values and tell me if they match within tolerance 0.01."""
    
    print("\n🚀 Starting Hybrid Orchestrator...\n")
    run_orchestrator(user_prompt)
