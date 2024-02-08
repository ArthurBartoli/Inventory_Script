import json

def unwrap_json(file_path):
    with open(file_path, 'r', encoding='utf-16') as json1_file:
        json_str = json1_file.read()
        return json.loads(json_str)
