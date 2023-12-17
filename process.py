import ai
import json
import main
def convert_string_to_json_file(json_string, output_file_path):
    try:
        # Parse the JSON-formatted string
        json_data = json.loads(json_string)

        # Write the JSON data to a file
        with open(output_file_path, 'w') as json_file:
            json.dump(json_data, json_file, indent=4)

        print(f'Successfully converted and saved JSON to {output_file_path}')
    except json.JSONDecodeError as e:
        print(f'Error decoding JSON string: {e}')
        
def convert_array_to_dict(array):
    result_dict = {}
    for index, item in enumerate(array, start=1):
        result_dict[index] = item
    return result_dict

def json_string_to_dict(json_string):
    try:
        # Parse the JSON-formatted string to a dictionary
        json_data = json.loads(json_string)
        return json_data
    except json.JSONDecodeError as e:
        print(f'Error decoding JSON string: {e}')
        return None
def to_array(excel):
    test = []
    for I in excel:
        temp = 'Extract the theme for each response to the question and put it into a json format that is structured like this {"question":"","Responses":[{"theme":"","response":""}]} the question is: ' + f'"{I}" the responses are: '
        responses = ''
        for J in excel[I]:
            responses += f'"{excel[I][J]}", '
        temp += responses
        test.append(temp)
    return(test)

def analyze(json):
    return 1
#test_dict = json_string_to_dict(ai.prompt('Extract the theme for each response to the question and put it into a json format that is structured like this {"question":"","Responses":[{"theme":"","response":""}]} the question is:' + f'"{}" the responses are: "Yes, sometimes because it covers half of your face. Especifically, when I have acne.", "As I stated about it hides my appearance", "It helps me show my confident-self.", "It doest help me since I dont have social anxiety and Im not self conscious.","It doesnt really help me with my social anxiety since it is just partly covers half of my face and it doesnt hide you from everyone. By self consciousness I think mask are a great help because it can cover my insecurities that i wanted to hide outside from people.","In terms of coping with social anxiety and Self-consciousness, facemask does not really help. Since it only covers a part of my face and doesnt really hide me from individuals.","As it covers half of my face, I know people cant judge me. Hence, it boost my self-confidence.","It helps me to hide my insecurities specially when Im in public places.","It can hide my nose and mouth which is my insecurities."')["choices"][0]["message"]["content"]))