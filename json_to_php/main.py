import json


def generate_php_file(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8-sig') as json_file, open(output_file, 'w', encoding='utf-8-sig') as php_file:
        data = json.load(json_file)
        php_file.write("<?php\n")
        for key, value in data.items():
            escaped_value = value.replace('\\', '\\\\').replace('"', '\\"') if value is not None else ""
            php_file.write(f'${key} = "{escaped_value}";\n')
        php_file.write("?>")


input_file = 'expression-list-ITA.json'
output_file = 'output.php'
generate_php_file(input_file, output_file)
