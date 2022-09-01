import json
import os
from create_document import create_document

# Open dynamic data for render
with open("dynamic_data.json") as dynamic_data_file:
    dynamic_data = json.load(dynamic_data_file)

# Open templates settings
with open('templates_settings.json', 'r') as templates_settings_file:
    templates_settings = json.load(templates_settings_file)

# Constants
TEMPLATES_DIRECTORY = templates_settings['templates_directory']
TEMPLATES_LIST = templates_settings['templates_list']
DESTINATION_DIRECTORY = 'generated/' + dynamic_data['object_name']

# Create destination directory
try:
    print('Создаём директорию: ' + DESTINATION_DIRECTORY)
    os.mkdir(DESTINATION_DIRECTORY)
except FileExistsError:
    print('Директория ' + DESTINATION_DIRECTORY + ' уже существует')

# Main process
for template in TEMPLATES_LIST:
    template_name = template[0]
    doc_name = template[1]
    doc_status = template[2]
    if doc_status == 1:
        print("Шаблон " + template_name + " включен. Поиск файла...")
        for entry in os.scandir(TEMPLATES_DIRECTORY):
            if entry.path.endswith(".docx") and entry.is_file():
                entry_name = os.path.splitext(entry.name)[0]
                if template_name == entry_name:
                    try:
                        print("Файл шаблона найден. Генерация...")
                        create_document(entry.path, dynamic_data,
                                        DESTINATION_DIRECTORY, doc_name)
                    except Exception as e:
                        print(e)
