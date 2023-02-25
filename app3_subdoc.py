# Возвращает поддокумент приложения В.
import io
import base64
from docx.shared import Inches, Mm


def app3_subdoc(data, doc):
    sd = doc.new_subdoc()
    table = sd.add_table(rows=0, cols=2)
    counter = 0
    for file_entry in data["КартинкиПриложенияВ"]:
        counter += 1
        if counter % 2 != 0:
            row = table.add_row()
        else:
            row = table.rows[len(table.rows) - 1]
        file_guid = file_entry["ИмяФайла"]
        file_extension = file_entry["Расширение"]
        file_base64 = file_entry["ДанныеФайла"]
        image = base64.b64decode(file_base64)
        filename = f'workdir/{file_guid}.{file_extension}'
        with open(filename, "wb") as file:
            file.write(image)

        # Добавляем картинку в файл.
        current_col = 0 if counter % 2 == 0 else 1

        row.cells[current_col].paragraphs[0].add_run().add_picture(
            filename, width = Mm(70), height = None)

    sd.add_paragraph("Тест!!!")
    return sd
