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

enabled_templates_list = [template for template in TEMPLATES_LIST if template['enabled']]

# Main process
for template in enabled_templates_list:
    template_name = template['input_name']
    doc_name = template['output_name']
    template_path = TEMPLATES_DIRECTORY + "/" + template_name + ".docx"
    try:
        create_document(template_path, dynamic_data,
                        DESTINATION_DIRECTORY, doc_name)
    except Exception as e:
        print("❌", e)
