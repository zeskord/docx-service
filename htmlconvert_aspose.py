import aspose.words as aw

# Load the document from disk
doc = aw.Document("workdir/generated_doc.docx")

# Enable export of fonts
options = aw.saving.HtmlSaveOptions()
options.export_font_resources = True
  
# Save the document as HTML
doc.save("workdir/output.html", options)