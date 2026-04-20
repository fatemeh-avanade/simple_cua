# gpt_ocr_test.py

import json
import base64
import re
from openai import AzureOpenAI
import os
from dotenv import load_dotenv
load_dotenv()

def safe_json_parse(text):
    """Parse JSON from text, handling markdown code blocks."""
    # Remove markdown code block if present
    text = re.sub(r'^```json\s*', '', text.strip())
    text = re.sub(r'\s*```$', '', text)
    return json.loads(text)

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

client = AzureOpenAI(
    api_key=OPENAI_API_KEY,
    api_version="2024-02-15-preview",
    azure_endpoint="https://fa-test-openai-instance-canada-east.openai.azure.com/"
)

# First test model with text only
response = client.chat.completions.create(
    model="fa-test-gpt-4o",
    messages=[
        {"role": "system", "content": "You are a helpful assistant."},
        {"role": "user", "content": "Hello from Python"}
    ],
)
print(response.choices[0].message.content)

# Now test image processing

data_path = r"C:\Users\fatemeh.torabi.asr\small_vsc_projects\quick_python_project\quick_python_project\hybrid_orchestration_demo\data"
image_path = os.path.join(data_path, 'terminal_test.png')

# Convert image to base64
with open(image_path, "rb") as image_file:
    image_data = base64.b64encode(image_file.read()).decode('utf-8')

response = client.chat.completions.create(
    model="fa-test-gpt-4o",
    messages=[
        {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": (
                        "Read the screenshot and extract the TOTAL value.\n"
                        "Return ONLY valid JSON like the following without extra formatting:\n"
                        "{ \"total\": 1000.0 }"
                    )
                },
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/png;base64,{image_data}"
                    }
                }
            ]
        }
    ]
)

print(response.choices[0].message.content)
parsed = safe_json_parse(response.choices[0].message.content)
print("Parsed total:", parsed["total"])