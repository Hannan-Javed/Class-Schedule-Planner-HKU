from PyInquirer import prompt

def list_menu_selector(field, prompt_message, choices):
    menu = prompt([
        {
            'type': 'list',
            'name': field,
            'message': prompt_message,
            'choices': choices,
        }
    ])
    return menu.get(field, None)