import requests


def prompt(message):
    headers = {
        'Authorization': 'Bearer pk-xhKjSwoKuugomeWXqSuCricvKUPYcoZRhBYJJWPxJLwCJEUr',
        'Content-Type': 'application/json',
    }

    json_data = {
        'model': 'pai-001-light-beta',
        'max_tokens': 2000,
        'messages': [
            {
                'role': 'system',
                'content': 'You are an helpful assistant.',
            },
            {
                'role': 'user',
                'content': message,
            },
        ],
    }

    response = requests.post('https://api.pawan.krd/v1/chat/completions', headers=headers, json=json_data)
    return response.json()

def test():
    return "1"