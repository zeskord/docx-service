# Возвращает поддокумент приложения В.

def app3_subdoc(data, doc):
    sd = doc.new_subdoc()
    # table = sd.add_table(rows=0, cols=2)
    for file_entry in data["КартинкиПриложенияВ"]:
        sd.add_paragraph(file_entry["ИмяФайла"])


    sd.add_paragraph("Тест!!!")
    return sd
