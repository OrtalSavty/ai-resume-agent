from flask import Flask, render_template_string, send_file, request
import json
import os
import webbrowser
import threading
import zipfile
from docx2pdf import convert
from generate_word import create_full_word_resume

try:
    import pythoncom
except ImportError:
    pass

app = Flask(__name__)

#  עיצוב ומבנה הממשק
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
    </style>
</head>
<body>
    <div class="container">
        <h1>מערכת חכמה לניהול והתאמת קורות חיים</h1>
        
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

        <div class="section">
            <h2>התאמה למשרה ספציפית</h2>
            <p>הקלידי את פרטי המשרה והמערכת תתאים את קורות החיים במיוחד עבורה.</p>
            <form action="/generate_tailored" method="post">
                <label>בחירת שפת קורות החיים:</label>
                <select name="lang">
                    <option value="he">עברית</option>
                    <option value="en">אנגלית</option>
                </select>

                <label>שם המשרה המבוקשת:</label>
                <input type="text" name="role" placeholder="למשל: מנתחת נתונים / מפתחת תוכנה" required>
                
                <label>תיאור המשרה:</label>
                <textarea name="jd" rows="4" placeholder="הדביקי לכאן את דרישות המשרה כדי שהמערכת תנתח אותן..."></textarea>
                
                <button class="btn" type="submit">התאם קורות חיים והורד קבצים</button>
            </form>
        </div>
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

@app.route('/generate_tailored', methods=['POST'])
def generate_tailored():
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
        
    lang = request.form.get('lang', 'he')
    role = request.form.get('role', 'Target_Role')
    
    active_master_file = 'master_resume_he.json' if lang == 'he' else 'master_resume.json'
    
    with open(active_master_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
    original_summary = data.get('professional_summary', '')
    
    # משנה את משפט הפתיחה בהתאם לשפה שנבחרה    
    if lang == 'he':
        data['professional_summary'] = f"מועמדת בעלת מוטיבציה גבוהה המכוונת למשרת {role}. " + original_summary
    else:
        data['professional_summary'] = f"Highly motivated candidate targeting the {role} position. " + original_summary
    
    tailored_file = f"resume_{role.replace(' ', '_')}.json"
    with open(tailored_file, 'w', encoding='utf-8') as tf:
        json.dump(data, tf, ensure_ascii=False, indent=2)
        
    output_docx = f"Resume_Tailored.docx"
    output_pdf = f"Resume_Tailored.pdf"
    output_zip = f"Resume_Tailored_Files.zip"
    
    create_full_word_resume(tailored_file, output_docx)
    try:
        convert(output_docx, output_pdf)
    except Exception as e:
        print("PDF Error:", e)
        
    package_files(output_zip, output_docx, output_pdf)
    return send_file(output_zip, as_attachment=True)

def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000")

if __name__ == '__main__':
    threading.Timer(1.0, open_browser).start()
    app.run(port=5000)