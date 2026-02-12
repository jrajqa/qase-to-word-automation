# Jay mataji!
import os
import requests
import re # For cleaning HTML tags
from dotenv import load_dotenv
from docxtpl import DocxTemplate
from jinja2 import Environment

# 1. Get the absolute path of the folder where main.py sits
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# 2. Join it with the subfolder and filename
# This creates a full path like /Users/ironman/word-document-automation/templates/template.docx
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "template.docx")
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "Generated_Test_Cases.docx")

load_dotenv()
API_TOKEN = os.getenv("QASE_TOKEN")
PROJECT_CODE = os.getenv("PROJECT_CODE")
SUITE_ID = 25  # Change to your Suite ID

def clean_html(raw_html):
    """Removes HTML tags like <p> from Qase data for Word."""
    if not raw_html: return ""
    clean_text = re.sub('<.*?>', '', raw_html)
    return clean_text

def fetch_qase_data():
    # Change the URL to this:
    url = f"https://api.qase.io/v1/case/{PROJECT_CODE}"
    headers = {"accept": "application/json", "Token": API_TOKEN}
    params = {"suite_id": SUITE_ID, "limit": 100}

    try:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        raw_cases = response.json()['result']['entities']
        
        # Clean the data before sending to Word
        for case in raw_cases:
            case['description'] = clean_html(case.get('description', ''))
            case['preconditions'] = clean_html(case.get('preconditions', ''))
            for step in case.get('steps', []):
                step['action'] = clean_html(step.get('action', ''))
                step['expected_result'] = clean_html(step.get('expected_result', ''))
        
        return raw_cases
    except Exception as e:
        print(f"Error: {e}")
        return None

if __name__ == "__main__":
    cases = fetch_qase_data()
    if cases:
        doc = DocxTemplate(TEMPLATE_PATH)
        
        # 1. Create a Jinja2 Environment
        jinja_env = Environment()
        
        # 2. Register the 'zfill' filter so the template can use it
        jinja_env.filters['zfill'] = lambda s, width: str(s).zfill(width)
        
        # 3. Pass the environment into the render function
        doc.render({"cases": cases}, jinja_env)
        doc.save("output/Generated_Test_Cases.docx")
        print(f"✅ Created document with {len(cases)} test cases!")
