# סוכן ה-AI להתאמה ודיוק של תקציר קורות החיים (OpenRouter / DeepSeek)

import os
from dotenv import load_dotenv
from openai import OpenAI

# טעינת המפתח מקובץ .env
load_dotenv(override=True)
api_key = os.getenv("DEEPSEEK_API_KEY")

# אתחול הלקוח מול OpenRouter עם הגבלת זמן של 20 שניות
client = None
if api_key:
    try:
        client = OpenAI(
            api_key=api_key,
            base_url="https://openrouter.ai/api/v1",
            timeout=20.0
        )
    except Exception as e:
        print("Warning: Failed to initialize OpenRouter Client:", e)


def call_llm(prompt: str, provider: str = None) -> str:
    """קריאה ישירה ל-DeepSeek דרך OpenRouter"""
    print("\n[AI Agent] שולח קריאה ל-OpenRouter (DeepSeek)...")

    if not client:
        raise RuntimeError("לקוח ה-AI לא אותחל. בדקי את המפתח ב-.env")

    try:
        response = client.chat.completions.create(
            model="deepseek/deepseek-chat",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7
        )
        print("[AI Agent] התקבלה תשובה בהצלחה תוך שניות!")
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"\n[AI Error]: {e}\n")
        raise e


def generate_tailored_summary(original_summary, role, jd, candidate_skills, lang='he', provider=None):
    lang_instruction = (
        "התקציר חייב להיכתב בעברית מקצועית, טבעית ורהוטה."
        if lang == 'he' else
        "The summary must be written in professional, high-level Executive English."
    )

    prompt = f"""
    אתה מומחה לכתיבת קורות חיים בכירים.
    
    משרת היעד: '{role}'.
    תיאור ודרישות המשרה:
    {jd}
    
    כישורי המועמדת:
    {candidate_skills}
    
    התקציר המקורי:
    "{original_summary}"
    
    הנחיות לביצוע:
    1. שכתב את פסקת התקציר כך שתבליט את הניסיון, המיומנויות והחוזקות של המועמדת בהתאמה ישירה לדרישות המשרה.
    2. כתיבה בגוף ראשון נסתר וישיר (ללא שימוש בכינויי גוף שלישי כמו "היא", "שלה", וללא לוכסנים מגדריים).
    3. שמור על אמינות מלאה בהתאם לכישורים הקיימים וללא המצאת ניסיון פיקטיבי.
    4. {lang_instruction}
    5. החזר אך ורק את פסקת התקציר המשוכתבת, ללא הקדמות, מרכאות או הסברים נוספים.
    """

    try:
        raw_result = call_llm(prompt, provider=provider)
        tailored = raw_result.replace('"', '').replace('```', '').strip()
        return tailored
    except Exception as e:
        print("Error during generate_tailored_summary:", e)
        return original_summary


def refine_tailored_summary(current_summary, feedback, role, jd, candidate_skills, lang='he', provider=None):
    lang_instruction = (
        "התקציר חייב להיכתב בעברית מקצועית וטבעית."
        if lang == 'he' else
        "The summary must be written in professional, high-level Executive English."
    )

    prompt = f"""
    אתה מומחה לכתיבת קורות חיים בכירים.
    
    קיבלת תקציר קיים עבור המשרה: '{role}'.
    תיאור ודרישות המשרה: {jd}
    כישורי המועמדת: {candidate_skills}
    
    התקציר הנוכחי:
    "{current_summary}"
    
    המשתמשת נתנה לך את המשוב וההנחיות הבאות לשיפור התקציר:
    "{feedback}"
    
    הנחיות לביצוע:
    1. שכתב ודייק את התקציר תוך יישום מדויק של הנחיות המשתמשת.
    2. כתיבה בגוף ראשון נסתר וישיר (ללא שימוש בכינויי גוף שלישי כמו "היא", "שלה", וללא לוכסנים מגדריים).
    3. שמור על אמינות בהתאם לכישורי המועמדת ולמשרת היעד.
    4. {lang_instruction}
    5. החזר אך ורק את פסקת התקציר המשוכתבת, ללא הקדמות, מרכאות או הסברים נוספים.
    """

    try:
        raw_result = call_llm(prompt, provider=provider)
        refined = raw_result.replace('"', '').replace('```', '').strip()
        return refined
    except Exception as e:
        print("Error during refine_tailored_summary:", e)
        return current_summary