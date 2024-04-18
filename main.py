from docxtpl import DocxTemplate

doc = DocxTemplate("template/Трудовой_договор_с_работником.docx")
context = {'name_surname': "<span style='color: yellow'>Жабинский В.В.</span>"}
doc.render(context)
doc.save("шаблон-final.docx")