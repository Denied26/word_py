from docxtpl import DocxTemplate
doc = DocxTemplate("Служебная записка.docx")
dict = {
    'to_position': 'Начальник отдела',
    'to_Fio': 'Некрасов В.И',
    'date': '26.10.2000',
    'number': '12345',
    'link': 'wii',
    'theme': 'Замена оборудования',
    'text': 'Требуется замена принтера в 123 кабинете',
    'from_Fio': 'Кучерев И.А',
    'from_position_Fio': 'Бухгалтер'
}
doc.render(dict)
doc.save("iw.docx")
