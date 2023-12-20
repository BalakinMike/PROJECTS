# подключаем библиотеки
import PySimpleGUI as sg
import os
import attestats_searching as ac

file_list_column = [
    [
        sg.Text("Папка с исходными файлами"),
        sg.In(size=(25, 1), enable_events=True, key="-FOLDER-"),
        sg.FolderBrowse('Обзор'),
        
    ],
    [sg.Text("Ввод поисковых параметров"),],
    [sg.Text("Номер документа   "), sg.In(size=(25, 1), enable_events=True, key="-DATA_N-"),],
    [sg.Text("Фамилия выпускника"), sg.In(size=(25, 1), enable_events=True, key="-DATA_LN-"),],
    [sg.Text("Город рождения    "), sg.In(size=(25, 1), enable_events=True, key="-DATA_Ct-"),],
    [sg.Button('Поиск', enable_events=True, key="-SEARCH-")],
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
            values=[], enable_events=True, size=(40, 20), key="-DATA_OUTPUT-"
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
            
        except:
            file_list = []
        fnames = [
            f
            for f in file_list
            if os.path.isfile(os.path.join(folder, f))
            and f.lower().endswith((".xlsx"))
        ]
        output_directory = (os.path.join(folder, "output"))
        
        try:
            os.makedirs(output_directory)
            print(f"Successfully created the directory: {output_directory}")
        except FileExistsError:
            print(f"The directory: {output_directory} already exists.")
        
        window["-FILE LIST-"].update(fnames)

    elif event == "-FILE LIST-":  # A file was chosen from the listbox
        try:
            filename = os.path.join(
                values["-FOLDER-"], values["-FILE LIST-"][0]
            )
            

        except:
            pass
    tasks = []
    if event == "-DATA_N-":
        t_task = 'Номер документа'
        task = values["-DATA_N-"]
        tasks = [t_task, task]
            
    elif event == "-DATA_LN-":
        t_task = 'Фамилия выпускника'
        task = values["-DATA_LN-"]
        tasks = [t_task, task]
                
    elif event == "-DATA_Ct-":
        t_task = 'Город рождения'
        task = values["-DATA_Ct-"]
        tasks = [t_task, task]
    print(tasks)
    if event == "-SEARCH-":
        print(tasks)
        if tasks != []:
            for filename in os.listdir(folder):
                if filename.endswith(".xlsx"):
                    ac.read_and_filter_excel(os.path.join(folder, filename),tasks[0], tasks[1])
                    print(tasks[0], tasks[1])
                else:
                    print(f"Skipping non-xlsx file: {filename}")
       
    window["-DATA_OUTPUT-"].update(tasks)
    
# закрываем окно и освобождаем используемые ресурсы
window.close()