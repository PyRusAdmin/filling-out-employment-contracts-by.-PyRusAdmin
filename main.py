from docxtpl import DocxTemplate

doc = DocxTemplate("шаблон.docx")
context = { 'director' : "И.И.Иванов"}
doc.render(context)
doc.save("шаблон-final.docx")