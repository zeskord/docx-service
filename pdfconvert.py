from docx2pdf import convert

# Основное действие.
if __name__ == '__main__':
    convert("workdir/generated_doc.docx", "workdir/output.pdf")