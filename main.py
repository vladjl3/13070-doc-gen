from docxtpl import DocxTemplate
import json
import os


def create_document(template_path: str, context: dict, folder: str, name: str):
    """
    Creates '.docx' document in path 'folder', with document name 'name', by template with jinja2 tags located in 'template_path'.
    'context' parameter is the dict object, with key-value of jinja2 tags.

    Args:
        template_path (str): path of .docx template with jinja2 tags
        context (dict): key-value dict of jinja2 tags
        folder (str): path of destination folder
        name (str): name of creating file
    Return: None
    """
    tpl = DocxTemplate(template_path)
    tpl.render(context)
    tpl.save(folder + "/" + name + ".docx")


# Open settings
with open("settings.json") as json_file:
    context_settings = json.load(json_file)

# Constants
TEMPLATES_DIRECTORY = context_settings['templates_directory']
DESTINATION_DIRECTORY = context_settings['object_name']

# Create destination directory
try:
    print('Создаём директорию: ' + DESTINATION_DIRECTORY)
    os.mkdir(DESTINATION_DIRECTORY)
except FileExistsError:
    print('Директория ' + DESTINATION_DIRECTORY +  ' уже существует')

# Open templates list
with open('templates_list.json', 'r') as filehandle:
    templates_list = json.load(filehandle)

# Main process
for template in templates_list:
    template_name = template[0]
    doc_name = template[1]
    doc_status = template[2]
    if doc_status == 1:
        print("Шаблон " + template_name + " включен. Поиск файла...")
        for entry in os.scandir(TEMPLATES_DIRECTORY):
            if entry.path.endswith(".docx") and entry.is_file():
                entry_name = os.path.splitext(entry.name)[0]
                if template_name == entry_name:
                    print("Файл шаблона " + template_name +
                          " найден. Генерация...")
                    create_document(entry.path, context_settings,
                                    DESTINATION_DIRECTORY, doc_name)
