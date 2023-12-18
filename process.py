import ai
import json
import main
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment
from io import BytesIO
import asyncio

def dict_to_excel(dictionary):
    """
    Convert a nested dictionary to a formatted Excel file.

    Parameters:
    - dictionary (dict): The nested dictionary to be converted.

    Returns:
    - BytesIO: BytesIO object containing the Excel file.
    """
    asyncio.sleep(5)
    # Initialize lists to store data for each column
    question_list, response_list, theme_list = [], [], []

    # Set to keep track of added questions
    added_questions = set()

    # Iterate through the dictionary
    for key, value in dictionary.items():
        # Extract data from the dictionary
        question = value['question']
        responses = value['Responses']

        # Append data to lists
        for response in responses:
            # Check if the question has been added
            if question in added_questions:
                question_list.append('')
            else:
                question_list.append(question)
                added_questions.add(question)

            response_list.append(response['response'])
            theme_list.append(response['theme'])

    # Create a DataFrame
    df = pd.DataFrame({
        'Question': question_list,
        'Response': response_list,
        'Theme': theme_list
    })

    # Create a Pandas Excel writer using XlsxWriter as the engine
    with BytesIO() as excel_buffer:
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            # Write the DataFrame to the Excel file
            df.to_excel(writer, sheet_name='Sheet1', index=False)

            # Access the XlsxWriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            # Set column width
            for column in worksheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

            # Format the cells
            for row in worksheet.iter_rows(min_row=2, max_col=3, max_row=worksheet.max_row):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

        # Save the Excel buffer
        excel_buffer.seek(0)
        return excel_buffer.getvalue()

        
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