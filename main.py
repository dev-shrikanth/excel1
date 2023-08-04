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


def write_json_list_to_excel(workbook, json_list, sheet_title, named_range):
    sheet = workbook.create_sheet(title=sheet_title)

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
    named_range_obj = DefinedName(name=named_range,  attr_text=named_range_ref)
    workbook.defined_names.add(named_range_obj)

    print(f"Named range '{named_range}' created for the data in '{sheet.title}' sheet.")


if __name__ == "__main__":
    workbook = openpyxl.Workbook()

    # Accept 5 API URLs
    api_urls = []
    for i in range(5):
        api_url = input(f"Enter API URL {i + 1}: ")
        api_urls.append(api_url)

    names = []
    for i in range(5):
        name = input(f"Enter the range name for URL {i+1}: ")
        names.append(name)

    for i, api_url in enumerate(api_urls, start=1):
        json_response = get_api_response(api_url)

        if json_response and isinstance(json_response, list):
            sheet_title = f"API_{i}"
            named_range = f"{names[i-1]}"
            if named_range is None:
                named_range = f"API_{i}"
            write_json_list_to_excel(workbook, json_response, sheet_title, named_range)
        else:
            print(f"Invalid API response for URL {api_url}. The API should return a list of dictionaries.")

    # Save the Excel file
    output_filename = "output.xlsx"
    workbook.save(output_filename)
    print(f"\nJSON data successfully written to {output_filename}")
    print("Named ranges created for the data from multiple API calls.")
