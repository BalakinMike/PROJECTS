# подключаем библиотеки
import PySimpleGUI as sg
import os

file_list_column = [
    [
        sg.Text("Папка с исходными файлами"),
        sg.In(size=(25, 1), enable_events=True, key="-FOLDER-"),
        sg.FolderBrowse('Обзор'),
        
    ],
    [sg.Text("Ввод поисковых параметров"),],
    [sg.Text("Номер документа"), sg.In(size=(25, 1), enable_events=True, key="-DATA_N-"),],
    [sg.Text("Фамилия выпускника"), sg.In(size=(25, 1), enable_events=True, key="-DATA_LN-"),],
    [sg.Text("Город рождения"), sg.In(size=(25, 1), enable_events=True, key="-DATA_Ct-"),],
    [sg.FolderBrowse('Поиск'),],
    [
        sg.Listbox(
            values=[], enable_events=True, size=(40, 20), key="-FILE LIST-"
        )
    ],
]

# For now will only show the name of the file that was chosen
output_viewer_column = [
    [sg.Text("Файлы с найденными сведениями:")],
    [
        sg.Listbox(
            values=[], enable_events=True, size=(40, 20), key="-FILE LIST-"
        )
    ],
]

# ----- Full layout -----
layout = [
    [
        sg.Column(file_list_column),
        sg.VSeperator(),
        sg.Column(output_viewer_column),
    ]
]

window = sg.Window("Поиск в базе аттестатов", layout)

while True:
    # получаем события, произошедшие в окне
    event, values = window.read()
    if event == "Exit" or event == sg.WIN_CLOSED:
        break
    # если нажали на кнопку
    if event == "-FOLDER-":
        folder = values["-FOLDER-"]
        try:
            # Get list of files in folder
            file_list = os.listdir(folder)
            print(folder)
        except:
            file_list = []
        fnames = [
            f
            for f in file_list
            if os.path.isfile(os.path.join(folder, f))
            and f.lower().endswith((".xlsx"))
        ]
        window["-FILE LIST-"].update(fnames)
    elif event == "-FILE LIST-":  # A file was chosen from the listbox
        try:
            filename = os.path.join(
                values["-FOLDER-"], values["-FILE LIST-"][0]
            )
            window["-TOUT-"].update(filename)
            window["-IMAGE-"].update(filename=filename)

        except:
            pass
# закрываем окно и освобождаем используемые ресурсы
window.close()