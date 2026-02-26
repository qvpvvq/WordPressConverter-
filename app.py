import mammoth
from flask import Flask, render_template, request
from bs4 import BeautifulSoup

app = Flask(__name__)

allowed_extensions = {'docx'}


def is_allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions


@app.route("/", methods=["GET", "POST"])
def index():
    html_output = ""
    error = None
    
    if request.method == "POST":
        file = request.files.get("word_file")

        if file and is_allowed_file(file.filename):
            try:
                style_map = "table => table.wp-table"
                result = mammoth.convert_to_html(file, style_map=style_map)
                raw_html = result.value

                soup = BeautifulSoup(raw_html, "html.parser")
                
                for tag in soup.find_all():
                    if tag.get_text(strip=True) and len(tag.contents) == 0:
                        tag.decompose()
                html_output = str(soup)
            except Exception as e:
                error = f"Ошибка при чтении файла {e}"
        else:
            error = "Выберите корректный файл .docx"

    return render_template("index.html", html_output=html_output, error=error)


if __name__ == "__main__":
    app.run(debug=True)