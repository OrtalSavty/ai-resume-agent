# מרכז את כל הלוגיקה, הפרומפטים והתקשורת מול מודל ה-AI

import os
import time
from dotenv import load_dotenv
from google import genai
from google.genai.errors import ServerError, ClientError

load_dotenv()

def generate_tailored_summary(original_summary, role, job_description, candidate_skills="", lang='he'):
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("לא נמצא GEMINI_API_KEY בקובץ ה-.env")
        
    client = genai.Client(api_key=api_key)
    
    if lang == 'he':
        lang_instruction = "בעברית מקצועית ותקנית לקורות חיים"
        role_guide = """
- איסור מוחלט על גוף שלישי: אלו קורות החיים שלי! אל תשתמש במילים 'היא', 'לה', 'שלה'. נסח ישירות: 'בעלת ניסיון...', 'משלבת שליטה ב-...'
- שמירה על לשון נקבה טבעית (סטודנטית, מפתחת), ללא לוכסנים מגדריים כלל (ללא 'סטודנט/ית').
- ללא אזכור מפורש של שם המשרה (כמו 'למשרת...' או 'מתאימה לתפקיד...').
"""
    else:
        lang_instruction = "in fluent, impactful resume English"
        role_guide = """
- Write in first-person resume style (implied 'I', e.g., 'Results-driven student with solid experience in...', NEVER use third-person like 'She is' or 'Her skills').
- Do not explicitly mention phrases like 'targeting the role of' or 'suitable for the position'.
- Highlight relevant tools and strengths organically.
"""

    skills_context = f"\nארגז הכלים והטכנולוגיות:\n{candidate_skills}\n" if candidate_skills else ""

    prompt = f"""
אתה יועץ קריירה מומחה. תפקידך לנסח מחדש את פסקת התקציר (Professional Summary) של קורות החיים שלי, כך שתדגיש את החוזקות המתאימות ביותר לתחום המשרה.

פרטי המשרה:
- תחום: {role}
- דרישות: {job_description}

הרקע שלי:
- תקציר נוכחי: "{original_summary}"
{skills_context}

הנחיות קריטיות לסגנון:
{role_guide}
- מבנה: פסקה אחת חזקה, קוהרנטית ורציפה (3-4 משפטים) {lang_instruction}.
- פלט נקי: החזר אך ורק את פסקת התקציר עצמה, ללא שום מירכאות, כותרות או הקדמות.
"""

    # שימוש במודל העדכני של גוגל
    model_name = 'gemini-3.8-flash'

    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt
            )
            return response.text.strip()
        except ServerError:
            time.sleep(2)
            continue
        except ClientError as e:
            # אם יש עומס רגעי, המתנה וניסיון חוזר
            if "429" in str(e) and attempt < 2:
                time.sleep(2)
                continue
            raise e
            
    raise RuntimeError("לא ניתן היה להפיק תקציר כרגע. נסי שוב בעוד רגע.")