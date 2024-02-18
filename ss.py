import PySimpleGUI as sg

data = [
    [1, 'Lorem ipsum dolor sit amet consectetur adipisicing elit. Molestias dolores consequatur voluptate sit incidunt! Ex id error, ratione temporibus deserunt unde doloribus magni aliquam eius eum magnam aut rem maxime!\
    Adipisci expedita, quibusdam aliquam, quo animi aperiam vel rem, recusandae fuga tenetur quidem sequi excepturi voluptas possimus? Quis maiores excepturi officia at corporis fuga in! Iusto consequatur et itaque doloribus.\
    Expedita eos, tenetur optio sint suscipit libero velit debitis error, eum quis cum excepturi et magnam iste a incidunt commodi iure! Laboriosam quis necessitatibus odit, impedit temporibus amet quisquam incidunt.'],
    [3, 4],
    [5, 6],
    [7, 8],
    [9, 10],
    [11, 12],
    [13, 14],
    # Добавьте свои данные сюда
]

# Определите заголовки столбцов
headers = ['Column 1', 'Column 2']

# Определите тему оформления (например, 'DarkBlue3')
sg.theme('DarkBlue3')

# Определите макет окна
layout = [
    [sg.Table(values=data, headings=headers, auto_size_columns=True,
              display_row_numbers=False, justification='right',
              num_rows=min(25, len(data)), key='-TABLE-')],
    [sg.Button('Show Full Text'), sg.Button('Exit')],
    [sg.Multiline(size=(60, 10), key='-FULLTEXT-', visible=False, autoscroll=True)]
]

# Создайте окно
window = sg.Window('Table Example with Scrolling', layout, resizable=True)

# Запустите цикл обработки событий
while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED or event == 'Exit':
        break
    elif event == 'Show Full Text':
        selected_row = values['-TABLE-'][0]
        full_text = data[selected_row][1]
        window['-FULLTEXT-'].update(value=full_text)
        window['-FULLTEXT-'].update(visible=True)

# Закройте окно
window.close()
