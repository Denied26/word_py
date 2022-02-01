from Note import Note

dict = {
    'to_position': 'Начальник отдела',
    'to_fio': 'Некрасов В.И',
    'date': '26.10.2000',
    'number': '123456',
    'link': 'wii',
    'theme': 'Замена оборудования',
    'to_io': 'Вечеслав Иванович',
    'text': 'Требуется замена принтера в 123 кабинете',
    'from_fio': 'Кучерев И.А',
    'from_position': 'Бухгалтер',
    'performer': 'Дронов К.С.',
    't_number': '1-234-5'
}
arr =[ {
    'to_position': 'Начальник отдела',
    'text': 'Требуется замена принтера в 123 кабинете',
    'from_fio': 'Кучерев И.А',
    'from_position': 'Бухгалтер',
    'performer': 'Дронов К.С.',
    },
    {

    'number': '123456',
    'link': 'wii',
    'theme': 'Замена оборудования',
    'from_position': 'Бухгалтер',
    'performer': 'Дронов К.С.',
    't_number': '1-234-5'
}
]
arr2 = [[1, 2, 3], [3, 4, 5], [6, 7, 8]]
its = Note(note_arr=dict, table_arr=arr)
its.makeWordFile()
its.addTableEasy(arr2)
its.addText('Проверка функционала добавления текста')