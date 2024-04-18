from docxtpl import DocxTemplate

doc = DocxTemplate("template/Трудовой_договор_с_работником.docx")
context = { 'director' : "И.И.Иванов"}
doc.render(context)
doc.save("шаблон-final.docx")