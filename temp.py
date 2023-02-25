from docxtpl import DocxTemplate
import json
from app3_subdoc import app3_subdoc

# Чтение входящих параметров.
data = json.load(open("workdir/data.json", encoding='utf-8-sig'))


doc = DocxTemplate("template.docx")


# application3 = doc.new_subdoc();



data["ПриложениеВ"] = app3_subdoc(data, doc)

doc.render(data)
doc.save("workdir/generated_doc.docx")