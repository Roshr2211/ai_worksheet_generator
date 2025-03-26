import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
import base64
import json
import re
from openai import OpenAI
import os
from datetime import datetime
import io
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# OpenRouter Configuration
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=st.secrets["OPENROUTER_API_KEY"]
)

def clean_json_content(raw_content):
    """
    Comprehensively clean and prepare JSON content for parsing.
    """
    # Remove LaTeX boxed formatting more robustly
    cleaned = re.sub(r'\\boxed\s*\{+', '{', raw_content)
    
    # Remove any LaTeX escape sequences that might break JSON
    cleaned = re.sub(r'\\[a-zA-Z]+', '', cleaned)  # Removes things like \frac, \textbf

    # Remove markdown code block markers
    cleaned = cleaned.replace('```json', '').replace('```', '').strip()
    
    # Ensure proper JSON formatting
    cleaned = re.sub(r',\s*}', '}', cleaned)  # Removes trailing commas before closing braces
    cleaned = re.sub(r'{\s*"', '{"', cleaned)  # Ensures no unnecessary spaces at the beginning
    
    return cleaned

def response(prompt_text, max_retries=3):
    progress_bar = st.progress(0)
    status_text = st.empty()

    for attempt in range(max_retries):
        try:
            status_text.info(f"Generating worksheet (Attempt {attempt + 1}/{max_retries})...")
            progress_bar.progress(min((attempt + 1) * 25, 75))
            
            logger.info(f"Attempt {attempt + 1}: Sending prompt")
            
            response = client.chat.completions.create(
                model="deepseek/deepseek-r1-zero:free",
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": "You are a helpful assistant that ONLY responds with valid JSON. No extra text, no markdown."},
                    {"role": "user", "content": prompt_text},
                ]
            )
            
            raw_content = response.choices[0].message.content
            logger.info(f"Raw response content: {raw_content}")
            
            status_text.info("Cleaning and parsing response...")
            progress_bar.progress(min((attempt + 1) * 40, 90))
            
            cleaned_content = clean_json_content(raw_content)
            logger.info(f"Cleaned content: {cleaned_content}")
            
            worksheet_data = json.loads(cleaned_content)
            
            if "worksheet" in worksheet_data and len(worksheet_data["worksheet"]) == 9:
                progress_bar.progress(100)
                status_text.success("Worksheet generated successfully!")
                
                # Create editable worksheet preview
                edited_worksheet = []
                st.markdown("### Edit Worksheet")
                for i, item in enumerate(worksheet_data["worksheet"], 1):
                    edited_item = st.text_input(f"Question {i}", value=item, key=f"worksheet_item_{i}")
                    edited_worksheet.append(edited_item)
                
                # Delay and clear progress indicators
                import time
                time.sleep(1)
                progress_bar.empty()
                status_text.empty()
                
                return edited_worksheet
            else:
                logger.warning(f"Invalid worksheet structure. Retrying. Content: {worksheet_data}")
                status_text.warning(f"Invalid worksheet. Retry attempt {attempt + 1}")
                
        except json.JSONDecodeError as e:
            logger.error(f"JSON Parsing Error: {e}")
            status_text.error(f"JSON Parsing Error on Attempt {attempt + 1}")
            logger.error(f"Problematic content: {raw_content}")
        
        except Exception as e:
            logger.error(f"Unexpected Error: {e}")
            status_text.error(f"Unexpected Error on Attempt {attempt + 1}")
        
    progress_bar.empty()
    st.error("Failed to generate worksheet after maximum retries. Please try again.")
    return None

def create_worksheet(subject_selection, topic):
    worksheet = response(prompt(subject_selection, topic))
    if worksheet:
        doc_download = word(worksheet, subject_selection)
        return doc_download
    return None
def prompt(subject, topic):
    return f"""Generate a worksheet in PURE JSON format with EXACTLY 9 elements. 
Ensure clear, educational language for each question:
{{
    "worksheet": [
        "{subject} Worksheet on {topic}",
        "First comprehension question about {topic}",
        "Second comprehension question about {topic}",
        "Third comprehension question about {topic}",
        "Fourth comprehension question about {topic}",
        "Fifth comprehension question about {topic}",
        "Multiple choice question 1 with a, b, c, d options",
        "Multiple choice question 2 with a, b, c, d options",
        "Close passage with blanks related to {topic}"
    ]
}}"""

# Rest of the code remains the same

def word(worksheet, subject):
    doc = Document()
    title = worksheet[0].replace(" ", "-").lower()

    sections = doc.sections
    for section in sections:
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri Light'
    font.size = Pt(12)
    font.color.rgb = RGBColor(40, 40, 40)
    
    current_date = datetime.now()
    date = current_date.strftime("%d.%m.%y")
    header_style = doc.styles["Header"]
    header_paragraph = doc.sections[0].header.paragraphs[0]
    header_paragraph.text = f"{subject}\t\t{date}"
    header_paragraph.style = header_style

    header_style = doc.styles["Heading1"] 
    header_paragraph = doc.add_paragraph()
    header_run = header_paragraph.add_run(worksheet[0])
    header_run.bold = True
    header_run.font.size = Pt(18) 

    for _ in range(2):
        doc.add_paragraph()

    for i, content in enumerate(worksheet[1:], start=1):
        subtitle_paragraph = doc.add_paragraph(f"{i}) {content}")
        subtitle_run = subtitle_paragraph.runs[0]
        subtitle_run.font.bold = True
        subtitle_run.font.size = Pt(14)
        subtitle_run.font.color.rgb = RGBColor(0, 28, 46)
        for _ in range(3):
            doc.add_paragraph()

    doc_name = f"{title}.docx"
  
    doc.save(doc_name)
    return doc_name

def create_worksheet(subject_selection, topic):
    worksheet = response(prompt(subject_selection, topic))
    if worksheet:
        doc_download = word(worksheet, subject_selection)
        return doc_download
    return None

st.set_page_config(
    page_title="Worksheet Generator",
    page_icon="🏫",
)

st.title("CASE")

subject_selection = st.selectbox(
    "Subject",
    ["Select a Subject"] + ["Enter Your Own Subject"] + ["Mathematics 🔢", "English 🇬🇧", "History 📜", "Geography 🌍", "Biology 🌿", "Chemistry 🧪", "Physics ⚙️", "Computer Science 💻", "Music 🎵", "Art 🎨", "Sports 🏃‍♂️", "Ethics 🤔", "Religion ⛪", "Politics 🗳️", "Economics 💹", "Philosophy 🤯", "Social Studies 👥", "Psychology 🧠", "Sociology 👩‍👩‍👧‍👦", "Foreign Language 🗣️", "Latin 🏛️", "Spanish 🇪🇸", "French 🇫🇷", "Italian 🇮🇹", "Russian 🇷🇺", "Chinese 🇨🇳", "Japanese 🇯🇵", "Korean 🇰🇷", "Arabic 🇸🇦", "Media Studies 📱"],
    key="subject_dropdown"
)

if subject_selection != "Select a Subject":
    if subject_selection == "Enter Your Own Subject":
        subject_selection = st.text_input("Enter Your Own Subject")[:-1]
        st.empty()

topic = st.text_input("Topic:")

create_button = st.button("Create Worksheet")

if topic and create_button:
    doc_download = create_worksheet(subject_selection, topic)
    
    if doc_download:
        bio = io.BytesIO()
        doc = Document(doc_download)
        doc.save(bio)

        st.success("Worksheet created successfully!")

        st.download_button(
            label="Click here to download",
            data=bio.getvalue(),
            file_name="Worksheet.docx",
            mime="docx",
            key="download_button",
            help="green"
        )
    else:
        st.error("Failed to create worksheet. Please try again.")