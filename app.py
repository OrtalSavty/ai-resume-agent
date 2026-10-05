# מנהל את שרת ה-Flask, הניתובים והממשק

from flask import Flask, render_template_string, send_file, request
import json
import os
import re
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

# תיקיות ונתיבי קבצים
RESUME_SOURCE_DIR = "Resume"
OUTPUT_DIR = "generated_resumes"
os.makedirs(OUTPUT_DIR, exist_ok=True)

MASTER_FILES = {
    'he': os.path.join(RESUME_SOURCE_DIR, "Hebrew_Resume.json"),
    'en': os.path.join(RESUME_SOURCE_DIR, "English_Resume.json")
}

def get_master_file(lang):
    file_path = MASTER_FILES.get(lang, MASTER_FILES['he'])
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"קובץ קורות החיים לא נמצא בנתיב: {os.path.abspath(file_path)}")
    return file_path

def smart_delete_item(data, search_text):
    """מאתר ומוחק פריט (פרויקט שלם, קורס, מיומנות או טקסט) בצורה חכמה וסלחנית"""
    clean_text = search_text.strip().rstrip(':').rstrip('.').strip()
    if not clean_text:
        return False
        
    deleted = False

    # 1. חיפוש ומחיקה של פרויקט שלם (שם ותיאור)
    if 'projects' in data and isinstance(data['projects'], list):
        initial_len = len(data['projects'])
        data['projects'] = [
            p for p in data['projects']
            if not (isinstance(p, dict) and (clean_text in p.get('name', '') or (len(clean_text) > 8 and clean_text in p.get('description', ''))))
        ]
        if len(data['projects']) < initial_len:
            return True

    # 2. חיפוש ומחיקה מתוך השכלה (notes / קורסים)
    if 'education' in data and isinstance(data['education'], list):
        for edu in data['education']:
            for k, v in edu.items():
                if isinstance(v, str) and (clean_text in v or search_text in v):
                    new_v = v.replace(search_text, "").replace(clean_text, "")
                    new_v = re.sub(r',\s*,', ',', new_v)
                    new_v = re.sub(r',\s*\.', '.', new_v)
                    new_v = re.sub(r'\.\s*,', '.', new_v)
                    new_v = re.sub(r'\s+', ' ', new_v).strip().rstrip(',').lstrip(',')
                    edu[k] = new_v
                    deleted = True
                elif isinstance(v, list):
                    initial_len = len(v)
                    edu[k] = [item for item in v if clean_text not in str(item)]
                    if len(edu[k]) < initial_len:
                        deleted = True

    # 3. חיפוש ומחיקה מתוך מיומנויות (skills)
    if 'skills' in data and isinstance(data['skills'], dict):
        for cat, items in data['skills'].items():
            if isinstance(items, list):
                initial_len = len(items)
                data['skills'][cat] = [it for it in items if clean_text not in str(it)]
                if len(data['skills'][cat]) < initial_len:
                    deleted = True
            elif isinstance(items, str) and (clean_text in items or search_text in items):
                new_it = items.replace(search_text, "").replace(clean_text, "")
                new_it = re.sub(r',\s*,', ',', new_it)
                new_it = re.sub(r'\s+', ' ', new_it).strip().rstrip(',').lstrip(',')
                data['skills'][cat] = new_it
                deleted = True

    # 4. חיפוש בשאר הטקסטים (תמצית וכדומה)
    for field in ['professional_summary']:
        if field in data and (clean_text in data[field] or search_text in data[field]):
            new_s = data[field].replace(search_text, "").replace(clean_text, "")
            new_s = re.sub(r'\s+', ' ', new_s).strip()
            data[field] = new_s
            deleted = True

    return deleted

def smart_replace_text(obj, find_text, replace_text):
    """החלפת טקסט גמישה וסורקת"""
    clean_find = find_text.strip().rstrip(':').strip()
    
    if isinstance(obj, str):
        if find_text in obj:
            return obj.replace(find_text, replace_text)
        elif clean_find and clean_find in obj:
            return obj.replace(clean_find, replace_text)
        return obj
    elif isinstance(obj, list):
        return [smart_replace_text(item, find_text, replace_text) for item in obj]
    elif isinstance(obj, dict):
        return {k: smart_replace_text(v, find_text, replace_text) for k, v in obj.items()}
    return obj

def add_course_smart(data, new_course, lang):
    for edu in data.get('education', []):
        if 'notes' in edu:
            if isinstance(edu['notes'], list):
                for idx, item in enumerate(edu['notes']):
                    if "קורס" in item or "course" in item.lower():
                        edu['notes'][idx] = item.rstrip().rstrip(',') + f", {new_course}"
                        return True
                edu['notes'].append(f"קורסים מרכזיים: {new_course}")
                return True
            elif isinstance(edu['notes'], str):
                if "קורס" in edu['notes'] or "course" in edu['notes'].lower():
                    edu['notes'] = edu['notes'].rstrip().rstrip(',') + f", {new_course}"
                    return True
                else:
                    edu['notes'] = edu['notes'].rstrip().rstrip(',') + f", קורסים מרכזיים: {new_course}"
                    return True

        for key in ['courses', 'relevant_courses', 'קורסים']:
            if key in edu:
                if isinstance(edu[key], list):
                    edu[key].append(new_course)
                    return True
                elif isinstance(edu[key], str):
                    edu[key] = edu[key].rstrip().rstrip(',') + f", {new_course}"
                    return True

    return False

# עיצוב ומבנה הממשק
HTML_PAGE = """
<!DOCTYPE html>
<html dir="rtl" lang="he">
<head>
    <meta charset="UTF-8">
    <title>מערכת ניהול קורות חיים</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #fdf2f8; margin: 0; padding: 40px; color: #333; }
        .container { max-width: 740px; margin: auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 15px rgba(219, 39, 119, 0.15); border-top: 5px solid #db2777; }
        h1 { color: #9d174d; text-align: center; margin-bottom: 25px; font-size: 26px; }
        h2 { color: #be185d; font-size: 19px; border-bottom: 2px solid #fbcfe8; padding-bottom: 10px; margin-top: 0; }
        .btn { background-color: #db2777; color: white; padding: 12px 24px; border: none; border-radius: 6px; cursor: pointer; font-size: 15px; font-weight: bold; width: 100%; transition: 0.3s; margin-top: 15px; }
        .btn:hover { background-color: #be185d; }
        .btn-secondary { background-color: #9d174d; }
        .btn-secondary:hover { background-color: #831843; }
        .section { background-color: #fff0f5; border: 1px solid #fbcfe8; padding: 22px; border-radius: 8px; margin-bottom: 22px; }
        label { font-weight: bold; display: block; margin-top: 12px; margin-bottom: 6px; color: #831843; font-size: 14px; }
        input, textarea, select { width: 100%; padding: 10px; border: 1px solid #f472b6; border-radius: 6px; box-sizing: border-box; font-family: inherit; font-size: 14px; }
        input:focus, textarea:focus, select:focus { outline: none; border-color: #db2777; box-shadow: 0 0 5px rgba(219, 39, 119, 0.3); }
        p { margin-top: 0; color: #555; font-size: 14px; line-height: 1.5; }
        .preview-box { background: white; border: 1px solid #f472b6; border-radius: 6px; padding: 14px; margin-top: 5px; line-height: 1.6; }
        .alert-success { background-color: #dcfce7; border: 1px solid #86efac; color: #166534; padding: 12px; border-radius: 6px; margin-bottom: 20px; font-weight: bold; text-align: center; }
        .alert-error { background-color: #fee2e2; border: 1px solid #fca5a5; color: #991b1b; padding: 12px; border-radius: 6px; margin-bottom: 20px; font-weight: bold; text-align: center; }
        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
    </style>
    <script>
        function updateAddFormFields() {
            var sec = document.getElementById('target_section').value;
            document.getElementById('div_general_text').style.display = (sec === 'courses' || sec === 'summary') ? 'block' : 'none';
            document.getElementById('div_skill_fields').style.display = (sec === 'skill') ? 'block' : 'none';
            document.getElementById('div_project_fields').style.display = (sec === 'project') ? 'block' : 'none';
            
            var lbl = document.getElementById('lbl_general');
            var inp = document.getElementById('inp_general');
            if (sec === 'courses') {
                lbl.innerText = 'שם הקורס להוספה:';
                inp.placeholder = 'למשל: סוכני AI (נלקח הסמסטר)';
            } else if (sec === 'summary') {
                lbl.innerText = 'המשפט להוספה לתמצית:';
                inp.placeholder = 'למשל: בעלת ראייה עסקית חדה ויכולת אנליטית מוכחת.';
            }
        }
    </script>
</head>
<body onload="updateAddFormFields()">
    <div class="container">
        <h1>מערכת חכמה לניהול והתאמת קורות חיים</h1>

        {% if success_message %}
            <div class="alert-success">{{ success_message }}</div>
        {% endif %}
        {% if error_message %}
            <div class="alert-error">{{ error_message }}</div>
        {% endif %}
        
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
                <button class="btn" type="submit">הורד קבצים (Word + PDF)</button>
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

        <!-- תצוגת השוואה והורדה -->
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
                <button class="btn btn-secondary" type="submit">הורד קורות חיים מעודכנים (Word + PDF)</button>
            </form>
        </div>
        {% endif %}

        <!-- אזור עדכון ועריכה קבועה בקורות החיים -->
        <div class="section" style="border-top: 3px solid #be185d;">
            <h2>עדכון ועריכה קבועה של קורות החיים (לתמיד)</h2>
            <p>שינויים שיתבצעו כאן יישמרו ישירות בתוך קובץ ה-JSON המקורי בתיקיית Resume.</p>

            <!-- כלי 1: חיפוש, החלפה ומחיקה -->
            <form action="/edit_resume" method="post" style="margin-bottom: 25px;">
                <input type="hidden" name="action_type" value="replace">
                <h3 style="color: #9d174d; font-size: 16px; margin-bottom: 5px;">1. החלפה או הסרה של טקסט / פרויקט</h3>
                
                <label>שפת הקובץ לעריכה:</label>
                <select name="lang">
                    <option value="he">עברית (Hebrew_Resume.json)</option>
                    <option value="en">אנגלית (English_Resume.json)</option>
                </select>

                <label>טקסט או שם פרויקט לחיפוש / מחיקה:</label>
                <input type="text" name="find_text" placeholder="למשל: מערכת ניתוח סנטימנט מבוססת AI" required>

                <label>טקסט חדש (להשאיר ריק כדי למחוק לחלוטין):</label>
                <input type="text" name="replace_text" placeholder="להשאיר ריק למחיקה מלאה">

                <button class="btn btn-secondary" type="submit">בצע החלפה / הסרה ושמור</button>
            </form>

            <hr style="border: 0; border-top: 1px dashed #f472b6; margin: 20px 0;">

            <!-- כלי 2: הוספת פריט חדש -->
            <form action="/edit_resume" method="post">
                <input type="hidden" name="action_type" value="append">
                <h3 style="color: #9d174d; font-size: 16px; margin-bottom: 5px;">2. הוספת פריט חדש לקורות החיים</h3>

                <div class="grid-2">
                    <div>
                        <label>שפת הקובץ:</label>
                        <select name="lang">
                            <option value="he">עברית (Hebrew_Resume.json)</option>
                            <option value="en">אנגלית (English_Resume.json)</option>
                        </select>
                    </div>
                    <div>
                        <label>מה ברצונך להוסיף:</label>
                        <select name="target_section" id="target_section" onchange="updateAddFormFields()">
                            <option value="courses">קורס אקדמי (הוספה לרשימת הקורסים)</option>
                            <option value="skill">מיומנות / כלי טכנולוגי (Skills)</option>
                            <option value="project">פרויקט חדש (New Project)</option>
                            <option value="summary">תמצית קורות החיים (סוף הפסקה)</option>
                        </select>
                    </div>
                </div>

                <!-- שדה לקורסים ותמצית -->
                <div id="div_general_text">
                    <label id="lbl_general">שם הקורס להוספה:</label>
                    <input type="text" name="new_text" id="inp_general" placeholder="למשל: סוכני AI (נלקח הסמסטר)">
                </div>

                <!-- שדות למיומנות -->
                <div id="div_skill_fields" style="display: none;">
                    <label>קטגוריית מיומנות (למשל: רקע טכני / שפות תכנות):</label>
                    <input type="text" name="skill_category" placeholder="למשל: רקע טכני (אם ריק - ישובץ בקטגוריה הראשית)">
                    
                    <label>המיומנות או הכלי להוספה:</label>
                    <input type="text" name="skill_item" placeholder="למשל: Docker, MongoDB">
                </div>

                <!-- שדות לפרויקט -->
                <div id="div_project_fields" style="display: none;">
                    <label>שם הפרויקט:</label>
                    <input type="text" name="project_name" placeholder="למשל: מערכת ניתוח סנטימנט מבוססת AI">
                    
                    <label>תיאור הפרויקט והטכנולוגיות ששימשו בפיתוח:</label>
                    <textarea name="project_desc" rows="3" placeholder="למשל: פיתוח מודל בפייתון לחיזוי מגמות ודשבורד מעקב ב-Power BI"></textarea>
                </div>

                <button class="btn" type="submit">הוסף לקורות החיים ושמור לתמיד</button>
            </form>
        </div>

    </div>
</body>
</html>
"""

def package_files(zip_path, docx_path, pdf_path):
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        if os.path.exists(docx_path):
            zipf.write(docx_path, arcname=os.path.basename(docx_path))
        if os.path.exists(pdf_path):
            zipf.write(pdf_path, arcname=os.path.basename(pdf_path))
    return zip_path

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
    active_master_file = get_master_file(lang)
    
    output_docx = os.path.join(OUTPUT_DIR, "Resume_General.docx")
    output_pdf = os.path.join(OUTPUT_DIR, "Resume_General.pdf")
    output_zip = os.path.join(OUTPUT_DIR, "Resume_General_Files.zip")
    
    create_full_word_resume(active_master_file, output_docx)
    try:
        convert(output_docx, output_pdf)
    except Exception as e:
        print("PDF Error:", e)
        
    package_files(output_zip, output_docx, output_pdf)
    return send_file(output_zip, as_attachment=True, max_age=0)

@app.route('/tailor', methods=['POST'])
def tailor_resume():
    lang = request.form.get('lang', 'he')
    role = request.form.get('role', '')
    jd = request.form.get('jd', '')
    
    active_master = get_master_file(lang)
    with open(active_master, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
    original_summary = data.get('professional_summary', '')
    
    candidate_skills = json.dumps(data.get('skills', {}), ensure_ascii=False)
    tailored_summary = generate_tailored_summary(original_summary, role, jd, candidate_skills, lang)
    
    data['professional_summary'] = tailored_summary
    tailored_filename = os.path.join(OUTPUT_DIR, f"tailored_resume_{lang}.json")
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

@app.route('/download_tailored', methods=['POST'])
def download_tailored():
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    lang = request.form.get('lang', 'he')
    role = request.form.get('role', 'Tailored')
    tailored_json = os.path.join(OUTPUT_DIR, f"tailored_resume_{lang}.json")
    
    safe_role = re.sub(r'[\\/*?:"<>|()]', "", role).strip().replace(" ", "_")
    if not safe_role:
        safe_role = "Tailored"

    output_docx = os.path.join(OUTPUT_DIR, "Resume_Tailored.docx")
    output_pdf = os.path.join(OUTPUT_DIR, "Resume_Tailored.pdf")
    zip_path = os.path.join(OUTPUT_DIR, f"Resume_{safe_role}_Files.zip")
    
    try:
        create_full_word_resume(tailored_json, output_docx)
        try:
            convert(output_docx, output_pdf)
        except Exception as e:
            print("PDF Conversion Warning:", e)
            
        package_files(zip_path, output_docx, output_pdf)
        return send_file(zip_path, as_attachment=True, max_age=0)
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

@app.route('/edit_resume', methods=['POST'])
def edit_resume():
    lang = request.form.get('lang', 'he')
    action_type = request.form.get('action_type')
    master_path = get_master_file(lang)

    with open(master_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    success_msg = None
    error_msg = None

    if action_type == 'replace':
        find_text = request.form.get('find_text', '').strip()
        replace_text = request.form.get('replace_text', '').strip()

        if not replace_text:
            # בקשת מחיקה מלאה
            was_deleted = smart_delete_item(data, find_text)
            if was_deleted:
                with open(master_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                success_msg = f"✓ הפריט '{find_text}' הוסר בהצלחה מקורות החיים ונשמר לצמיתות!"
            else:
                error_msg = f"הטקסט '{find_text}' לא נמצא בקובץ. ודאי שהוא נכתב במדויק."
        else:
            # בקשת החלפה
            updated_data = smart_replace_text(data, find_text, replace_text)
            if updated_data == data:
                error_msg = f"הטקסט '{find_text}' לא נמצא לצורך החלפה."
            else:
                with open(master_path, 'w', encoding='utf-8') as f:
                    json.dump(updated_data, f, ensure_ascii=False, indent=2)
                success_msg = f"✓ הטקסט '{find_text}' הוחלף בהצלחה ב-'{replace_text}' ונשמר לצמיתות!"

    elif action_type == 'append':
        target_section = request.form.get('target_section')

        if target_section == 'courses':
            new_text = request.form.get('new_text', '').strip()
            if not new_text:
                error_msg = "יש להזין שם קורס להוספה."
            elif add_course_smart(data, new_text, lang):
                with open(master_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                success_msg = f"✓ הקורס '{new_text}' נוסף בהצלחה לרשימת הקורסים ונשמר לצמיתות!"
            else:
                error_msg = "לא נמצא שדה קורסים מתאים בקובץ."

        elif target_section == 'summary':
            new_text = request.form.get('new_text', '').strip()
            if not new_text:
                error_msg = "יש להזין משפט להוספה לתמצית."
            else:
                data['professional_summary'] = data.get('professional_summary', '').rstrip() + f" {new_text}"
                with open(master_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                success_msg = f"✓ המשפט נוסף בהצלחה לתמצית ונשמר לצמיתות!"

        elif target_section == 'skill':
            skill_cat = request.form.get('skill_category', '').strip()
            skill_item = request.form.get('skill_item', '').strip()

            if not skill_item:
                error_msg = "יש להזין את שם המיומנות להוספה."
            else:
                if 'skills' not in data:
                    data['skills'] = {}
                skills_obj = data['skills']

                if isinstance(skills_obj, dict):
                    if not skill_cat:
                        skill_cat = list(skills_obj.keys())[0] if skills_obj else ("רקע טכני" if lang == 'he' else "Tools")
                    
                    if skill_cat in skills_obj:
                        val = skills_obj[skill_cat]
                        if isinstance(val, list):
                            val.append(skill_item)
                        elif isinstance(val, str):
                            skills_obj[skill_cat] = val.rstrip().rstrip(',') + f", {skill_item}"
                    else:
                        skills_obj[skill_cat] = [skill_item]
                elif isinstance(skills_obj, list):
                    skills_obj.append(skill_item)

                with open(master_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                success_msg = f"✓ המיומנות '{skill_item}' נוספה בהצלחה ונשמרה לצמיתות!"

        elif target_section == 'project':
            p_name = request.form.get('project_name', '').strip()
            p_desc = request.form.get('project_desc', '').strip()

            if not p_name:
                error_msg = "חובה להזין שם פרויקט."
            else:
                if 'projects' not in data:
                    data['projects'] = []
                data['projects'].append({"name": p_name, "description": p_desc})

                with open(master_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                success_msg = f"✓ הפרויקט '{p_name}' נוסף בהצלחה לרשימת הפרויקטים ונשמר לצמיתות!"

    return render_template_string(HTML_PAGE, success_message=success_msg, error_message=error_msg)

def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000")

if __name__ == '__main__':
    threading.Timer(1.0, open_browser).start()
    app.run(port=5000)