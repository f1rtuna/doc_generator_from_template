from docx import Document
document = Document()
paragraph = document.add_paragraph('Lorem ipsum dolor sit amet.')
document.save('test.docx')

#https://theautomatic.net/2019/10/14/how-to-read-word-documents-with-python/