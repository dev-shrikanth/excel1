import requests
import openpyxl
from openpyxl.workbook.defined_name import DefinedName


def get_api_response(api_url):
    response = requests.get(api_url)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Failed to get API response. Status Code: {response.status_code}")
        return None


def write_json_list_to_excel(json_list, output_filename, named_range):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write headers in the worksheet
    headers = list(json_list[0].keys())
    for col_num, header in enumerate(headers, start=1):
        sheet.cell(row=1, column=col_num, value=header)

    # Write data to the worksheet
    for row_num, item in enumerate(json_list, start=2):
        for col_num, header in enumerate(headers, start=1):
            sheet.cell(row=row_num, column=col_num, value=item.get(header, ""))

    # Create a named range
    named_range_ref = f'{sheet.title}!$A$1:${openpyxl.utils.get_column_letter(len(headers))}${len(json_list) + 1}'
    # named_range_obj = DefinedName(name=named_range, localSheetId=0, formula=named_range_ref)
    named_range_obj = DefinedName(name=named_range, localSheetId=0, attr_text=named_range_ref)
    # workbook.defined_names.definedName.append(named_range_obj)
    workbook.defined_names.add(named_range_obj)

    # Save the Excel file
    workbook.save(output_filename)
    print(f"JSON data successfully written to {output_filename}")
    print(f"Named range '{named_range}' created for the data.")


if __name__ == "__main__":
    api_url = "https://jsonplaceholder.typicode.com/posts"
    json_response = get_api_response(api_url)

    if json_response and isinstance(json_response, list):
        output_filename = "output.xlsx"
        named_range = input("Enter the name for the named range: ")
        write_json_list_to_excel(json_response, output_filename, named_range)
    else:
        print("Invalid API response. The API should return a list of dictionaries.")
