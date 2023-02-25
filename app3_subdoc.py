# Возвращает поддокумент приложения В.
import io, base64

def app3_subdoc(data, doc):
    sd = doc.new_subdoc()
    table = sd.add_table(rows=0, cols=2)
    for file_entry in data["КартинкиПриложенияВ"]:
        row = table.add_row()
        file_guid = file_entry["ИмяФайла"]
        file_extension = file_entry["Расширение"]
        row.cells[0].text = file_guid
        file_base64 = file_entry["ДанныеФайла"]
        bytesio_object = io.BytesIO(base64.decodebytes(bytes(file_base64, "utf-8")))
        with open(f'workdir/{file_guid}.{file_extension}', "w+") as f:
            (bytesio_object.getbuffer())
            
        

        row.cells[1].text = file_entry["ДанныеФайла"]
        # sd.add_paragraph(file_entry["ИмяФайла"])


    sd.add_paragraph("Тест!!!")
    return sd
