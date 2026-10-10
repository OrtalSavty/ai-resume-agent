import json
import os
from dotenv import load_dotenv
from google import genai
from google.genai import types


def read_resume(filepath='master_resume.json'):
    
   #קורא את קובץ קורות החיים המקורי   
    with open(filepath, 'r', encoding='utf-8') as file:
        return json.load(file)

def save_resume(data, filepath='updated_resume.json'):
    #שומר את קורות החיים המעודכנים לקובץ חדש   
    with open(filepath, 'w', encoding='utf-8') as file:
        # ensure_ascii=False: שומר על תקינות עברית אם תהיה
        # indent=4: מסדר את ה-JSON שיהיה קריא ויפה לעין ולא בשורה אחת
        json.dump(data, file, ensure_ascii=False, indent=4)
    print(f" קורות החיים המעודכנים נשמרו בהצלחה בקובץ: {filepath}")


def real_ai_agent(resume_data, new_info):

    print("\n--- מתחבר ל-Gemini API ---")
    
    # טעינת המפתח מקובץ ה-.env שיצרנו   
    load_dotenv()
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        print(" שגיאה: לא נמצא מפתח API בקובץ .env")
        return resume_data
        
    # אתחול הלקוח של גוגל   
    client = genai.Client(api_key=api_key)
    
    print(f"המידע שהמשתמש ביקש להוסיף:\n'{new_info}'\n")
    print("מעבד את הבקשה. זה עשוי לקחת כמה שניות...")
    
    # כאן אנחנו בונים את הפרומפט   
    # אנחנו מבקשים ממנו להחזיר לנו JSON תקין ולא סתם טקסט רגיל    
    prompt = f"""
    You are an AI resume updating agent. 
    Below is a user's current resume in JSON format.
    The user wants to add this new information: "{new_info}"
    
    Your task:
    1. Read the JSON.
    2. Understand where the new information belongs (e.g., 'skills', 'experience', 'projects', 'education').
    3. Update the JSON structure by integrating the new information contextually.
    4. Return ONLY the updated JSON. Do not write any markdown code blocks or explanations. Just pure JSON.
    
    Current Resume JSON:
    {json.dumps(resume_data, ensure_ascii=False)}
    """
    
    # שליחת הבקשה למודל   
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=prompt,
        config=types.GenerateContentConfig(
            #  נגדיר לו בכוח שהתשובה חייבת להיות בפורמט JSON            
            response_mime_type="application/json",
        )
    )
    
    # ממירים את התשובה שהוא החזיר (שהיא טקסט) בחזרה למילון פייתון    
    try:
        updated_resume = json.loads(response.text)
        print(" הסוכן סיים לעבד בהצלחה")
        return updated_resume
    except json.JSONDecodeError:
        print(" שגיאה: המודל לא החזיר JSON תקין. מחזיר את המקור")
        return resume_data
    
    
# כאן מתחילה הריצה של התוכנית
if __name__ == "__main__":
    # טעינת קורות החיים המקוריים    
    my_resume = read_resume('master_resume.json')
    
    # המידע החדש שאנחנו רוצים להוסיף    
    new_info_from_user = "לקחתי לאחרונה קורס על ארכיטקטורת תוכנה ולמדתי לשלב מודלי AI של Gemini לתוך אפליקציות פייתון"    
    # הפעלת הסוכן המדומה    
    new_resume = real_ai_agent(my_resume, new_info_from_user)
    
    # שמירת קורות החיים המעודכנים לקובץ חדש    
    save_resume(new_resume, 'resume_version_with_docker.json')
    
    
    
    
    
    
    
    
    