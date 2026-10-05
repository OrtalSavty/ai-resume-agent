# מנהל את שרת ה-Flask, הניתובים והממשק

from flask import Flask, render_template_string, send_file, request
import json
import os
import webbrowser
import threading
import zipfile
from docx2pdf import convert
from generate_word import create_full_word_resume
from ai_agent import generate_tailored_summary

try:
    import pythoncom
except ImportError:
    pass

app = Flask(__name__)

# עיצוב ומבנה הממשק
HTML_PAGE = """
<!DOCTYPE html>
<html dir="rtl" lang="he">
<head>
    <meta charset="UTF-8">
    <title>מערכת ניהול קורות חיים</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #fdf2f8; margin: 0; padding: 40px; color: #333; }
        .container { max-width: 700px; margin: auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 15px rgba(219, 39, 119, 0.15); border-top: 5px solid #db2777; }
        h1 { color: #9d174d; text-align: center; margin-bottom: 30px; font-size: 26px; }
        h2 { color: #be185d; font-size: 19px; border-bottom: 2px solid #fbcfe8; padding-bottom: 10px; margin-top: 0; }
        .btn { background-color: #db2777; color: white; padding: 12px 24px; border: none; border-radius: 6px; cursor: pointer; font-size: 16px; font-weight: bold; width: 100%; transition: 0.3s; margin-top: 20px; }
        .btn:hover { background-color: #be185d; }
        .section { background-color: #fff0f5; border: 1px solid #fbcfe8; padding: 25px; border-radius: 8px; margin-bottom: 25px; }
        label { font-weight: bold; display: block; margin-top: 15px; margin-bottom: 8px; color: #831843; }
        input, textarea, select { width: 100%; padding: 10px; border: 1px solid #f472b6; border-radius: 6px; box-sizing: border-box; font-family: inherit; font-size: 14px; }
        input:focus, textarea:focus, select:focus { outline: none; border-color: #db2777; box-shadow: 0 0 5px rgba(219, 39, 119, 0.3); }
        p { margin-top: 0; color: #555; }
        .preview-box { background: white; border: 1px solid #f472b6; border-radius: 6px; padding: 14px; margin-top: 5px; line-height: 1.6; }
    </style>
</head>
<body>
    <div class="container">
        <h1>מערכת חכמה לניהול והתאמת קורות חיים</h1>
        
        <!-- אזור הפקת קורות חיים כלליים -->
        <div class="section">
            <h2>הפקת קורות חיים כלליים</h2>
            <p>הורדת גרסת הבסיס המלאה כפי שהיא שמורה במערכת.</p>
            <form action="/generate_general" method="post">
                <label>בחירת שפת קורות החיים:</label>
                <select name="lang">
                    <option value="he">עברית</option>
                    <option value="en">אנגלית</option>
                </select>
                <button class="btn" type="submit">הורד קבצים</button>
            </form>
        </div>

        <!-- אזור התאמה למשרה -->
        <div class="section">
            <h2>התאמה למשרה ספציפית</h2>
            <p>הקלידי את פרטי המשרה והסוכן ישכתב את התקציר בהתאם לדרישות.</p>
            <form action="/tailor" method="post">
                <label>בחירת שפת קורות החיים:</label>
                <select name="lang">
                    <option value="he" {% if lang == 'he' %}selected{% endif %}>עברית</option>
                    <option value="en" {% if lang == 'en' %}selected{% endif %}>אנגלית</option>
                </select>

                <label>שם המשרה המבוקשת:</label>
                <input type="text" name="role" value="{{ role or '' }}" placeholder="למשל: מנתחת נתונים / מפתחת תוכנה" required>
                
                <label>תיאור המשרה ודרישות:</label>
                <textarea name="jd" rows="4" placeholder="הדביקי לכאן את דרישות המשרה כדי שהסוכן ינתח אותן...">{{ jd or '' }}</textarea>
                
                <button class="btn" type="submit">התאם תקציר והצג השוואה</button>
            </form>
        </div>

        <!-- בלוק הנראות: תצוגת השוואה והורדה (מופיע רק אחרי שהסוכן רץ) -->
        {% if show_preview %}
        <div class="section" style="border: 2px solid #db2777; background-color: #fff;">
            <h2>השוואת תוצאות הסוכן עבור: {{ role }}</h2>
            
            <label>התקציר המקורי:</label>
            <div class="preview-box" style="background-color: #fdf2f8; color: #666; font-size: 13.5px;">
                {{ original_summary }}
            </div>
            
            <label>התקציר המותאם החדש:</label>
            <div class="preview-box" style="background-color: #fce7f3; color: #111; font-weight: 500; border-right: 4px solid #db2777;">
                {{ tailored_summary }}
            </div>
            
            <form action="/download_tailored" method="post">
                <input type="hidden" name="lang" value="{{ lang }}">
                <input type="hidden" name="role" value="{{ role }}">
                <button class="btn" style="background-color: #9d174d;" type="submit">הורד קורות חיים מעודכנים (Word + PDF)</button>
            </form>
        </div>
        {% endif %}

    </div>
</body>
</html>
"""

def package_files(zip_name, docx_name, pdf_name):
    with zipfile.ZipFile(zip_name, 'w') as zipf:
        if os.path.exists(docx_name):
            zipf.write(docx_name)
        if os.path.exists(pdf_name):
            zipf.write(pdf_name)
    return zip_name

@app.route('/')
def home():
    return render_template_string(HTML_PAGE)

@app.route('/generate_general', methods=['POST'])
def generate_general():
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
        
    lang = request.form.get('lang', 'he')
    active_master_file = 'master_resume_he.json' if lang == 'he' else 'master_resume.json'
    
    output_docx = "Resume_General.docx"
    output_pdf = "Resume_General.pdf"
    output_zip = "Resume_General_Files.zip"
    
    create_full_word_resume(active_master_file, output_docx)
    try:
        convert(output_docx, output_pdf)
    except Exception as e:
        print("PDF Error:", e)
        
    package_files(output_zip, output_docx, output_pdf)
    return send_file(output_zip, as_attachment=True)

# ניתוב 1: יצירת התקציר המותאם והצגתו בממשק
@app.route('/tailor', methods=['POST'])
def tailor_resume():
    lang = request.form.get('lang', 'he')
    role = request.form.get('role', '')
    jd = request.form.get('jd', '')
    
    active_master = 'master_resume_he.json' if lang == 'he' else 'master_resume.json'
    with open(active_master, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
    original_summary = data.get('professional_summary', '')
    
    # שליפת המיומנויות והעברתן לסוכן
    candidate_skills = json.dumps(data.get('skills', {}), ensure_ascii=False)
    tailored_summary = generate_tailored_summary(original_summary, role, jd, candidate_skills, lang)
        
        
    # שמירת נתוני המשרה המותאמים לקובץ זמני ייעודי
    data['professional_summary'] = tailored_summary
    tailored_filename = f"tailored_resume_{lang}.json"
    with open(tailored_filename, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        
    return render_template_string(
        HTML_PAGE,
        original_summary=original_summary,
        tailored_summary=tailored_summary,
        lang=lang,
        role=role,
        jd=jd,
        show_preview=True
    )


import re

# ניתוב 2: הורדת קובצי הוורד וה-PDF של הגרסה המותאמת
@app.route('/download_tailored', methods=['POST'])
def download_tailored():
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    lang = request.form.get('lang', 'he')
    role = request.form.get('role', 'Tailored')
    tailored_json = f"tailored_resume_{lang}.json"
    
    # ניקוי תווים שאסורים בשמות קבצים בווינדוס (כמו סלאשים, נקודותיים וסוגריים)
    safe_role = re.sub(r'[\\/*?:"<>|()]', "", role).strip().replace(" ", "_")
    if not safe_role:
        safe_role = "Tailored"

    output_docx = "Resume_Tailored.docx"
    output_pdf = "Resume_Tailored.pdf"
    zip_filename = f"Resume_{safe_role}_Files.zip"
    
    try:
        create_full_word_resume(tailored_json, output_docx)
        try:
            convert(output_docx, output_pdf)
        except Exception as e:
            print("PDF Conversion Warning:", e)
            
        package_files(zip_filename, output_docx, output_pdf)
        return send_file(zip_filename, as_attachment=True)
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        
        

def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000")

if __name__ == '__main__':
    threading.Timer(1.0, open_browser).start()
    app.run(port=5000)