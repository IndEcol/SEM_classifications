import re
import openpyxl
import yaml


def parse_xlsx_and_save_as_yaml(xlsx_file):
    wb = openpyxl.load_workbook(xlsx_file)
    listSheet = [sheet for sheet in wb.sheetnames if re.match(r"CM_[^_]*_Definition", sheet)]
    print(listSheet)
    content = {}

    for sheet_name in listSheet:
        sheet_data = {
            "classification_info": {},
            "metadata": {},
            "classification_items_description": {}
        }
        is_metadata = False
        name = "_".join(sheet_name.split("_")[0:2])

        for row in wb[name + "_Definition"]:
            if row[0].value == "Column" or not row[0].value:  # TODO: change
                continue
            if is_metadata:
                sheet_data["metadata"][row[0].value] = row[1].value
                continue
            if row[0].value == "Metadata":
                is_metadata = True
                continue
            sheet_data["classification_info"][row[0].value] = row[1].value

        titles = []
        for row in wb[name + "_Items"]:
            if row[0].value == "id" or not row[0].value:
                titles = [cell.value for cell in row]
                continue
            for i, cell in enumerate(row[3:]):
                if cell.value:
                    dict_name = "classification_items_" + titles[i + 3]
                    if dict_name not in sheet_data.keys():
                        sheet_data[dict_name] = {}
                    sheet_data[dict_name][row[0].value] = cell.value
            content[sheet_name] = sheet_data
        print(sheet_data)
        with open('CIRCOMOD_Project_Wide_Classifications' + name + '.yaml', 'w') as file:
            yaml.dump(sheet_data, file, sort_keys=False)


# Usage example
if __name__ == "__main__":
    XLSX_FILE = 'CIRCOMOD_Project_Wide_Classifications.xlsx'
    parse_xlsx_and_save_as_yaml(XLSX_FILE)
