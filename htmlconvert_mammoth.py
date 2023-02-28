import mammoth
import base64

def convert_image(image):
    print("conv")
    with image.open() as image_bytes:
        encoded_src = base64.b64encode(image_bytes.read()).decode("ascii")

    return {
        "src": "data:{0};base64,{1}".format(image.content_type, encoded_src)
    }

with open("workdir/generated_doc.docx", "rb") as docx_file:
    result = mammoth.convert_to_html(docx_file, convert_image=mammoth.images.img_element(convert_image))
    html = result.value
    with open('workdir/output.html', 'w') as html_file:
        html_file.write(html)
