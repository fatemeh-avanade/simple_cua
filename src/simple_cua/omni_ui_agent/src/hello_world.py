
import os, sys, json, re
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import pandas as pd
from openai import OpenAI

# # --- env & client ---
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

ORG_HOST = os.getenv("CRM_HOST", "yourorg.crm.dynamics.com")
OPPORTUNITIES_URL = f"https://{ORG_HOST}/main.aspx?app=d365default&pagetype=entitylist&etn=opportunity"
USER_DATA_DIR = os.getenv("USER_DATA_DIR", str(Path.cwd() / ".pw-profile"))



if __name__ == "__main__":
    print("Hello world")
    print("Environment variables loaded:")
    print(f"CRM_HOST: {ORG_HOST}")
    print(f"USER_DATA_DIR: {USER_DATA_DIR}")
        