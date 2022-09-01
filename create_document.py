from docxtpl import DocxTemplate


def create_document(template_path: str, context: dict, folder: str, name: str):
    """
    Creates '.docx' document in path 'folder', with document name 'name',
    by template with jinja2 tags located in 'template_path'.
    'context' parameter is a dict object, with key-value of jinja2 tags.

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
    print("✅" + name + " успешно сгенерирован.")
