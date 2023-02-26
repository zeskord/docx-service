# Возвращает поддокумент приложения В.
import io
import base64
from docx.shared import Inches, Mm, Pt
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH


def app3_subdoc(data, doc):
    sd = doc.new_subdoc()
    
    picture_group_counter = 0
    for picture_group in data["КартинкиПриложенияВ"]:
        
        picture_group_counter += 1

        table = sd.add_table(rows=0, cols=2)

        counter = 0
        for file_entry in picture_group["ДанныеФайлов"]:
            counter += 1
            if counter % 2 != 0:
                row_pictures = table.add_row()
                row_pictures.height = Mm(100)
                row_descriptions = table.add_row()
            else:
                row_pictures = table.rows[len(table.rows) - 2]
                row_descriptions = table.rows[len(table.rows) - 1]
            file_guid = file_entry["ИмяФайла"]
            file_extension = file_entry["Расширение"]
            file_base64 = file_entry["ДанныеФайла"]
            image = base64.b64decode(file_base64)
            filename = f'workdir/{file_guid}.{file_extension}'
            with open(filename, "wb") as file:
                file.write(image)

            # Добавляем картинку в файл.
            current_col = 0 if counter % 2 == 0 else 1
            cell_pictures = row_pictures.cells[current_col]
            cell_pictures.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
            picture_paaragraph = cell_pictures.paragraphs[0]
            picture_paaragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            picture_paaragraph.paragraph_format.space_after = Pt(0)
            run = picture_paaragraph.add_run()
            picture = run.add_picture(filename, width=Mm(75), height=None)

            cell_descriptions = row_descriptions.cells[current_col]
            description = file_entry["Описание"]
            par_description = cell_descriptions.paragraphs[0]
            par_description.alignment = WD_ALIGN_PARAGRAPH.CENTER
            description_run = par_description.add_run(description)
            description_run.font.size = Pt(12)
        
        sd.add_paragraph(f'Рисунок {picture_group_counter}: {picture_group["ОписаниеРисунка"]}')


    sd.add_paragraph("Конец!!!")
    return sd
