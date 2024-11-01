from flask import Flask, request, send_file, render_template
from pptx import Presentation
import pandas as pd
import requests
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def upload_form():
    return render_template('upload.html')

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    df = pd.read_excel(file)

    # יצירת מצגת
    prs = Presentation()

    for index, row in df.iterrows():
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        
        # הוספת כותרת
        title = slide.shapes.title
        title.text = str(row['כותרת'])

        # הורדת תמונה אם יש כתובת URL
        img_url = row.get('URL לתמונה')
        if pd.notna(img_url):
            response = requests.get(img_url)
            img = BytesIO(response.content)
            slide.shapes.add_picture(img, left=0, top=150, width=prs.slide_width * 0.8)

        # הוספת טקסטים נוספים לפי עמודות אחרות
        # ...

    # שמירת קובץ PowerPoint
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="output.pptx")

if __name__ == "__main__":
    app.run(debug=True)
