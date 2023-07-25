import re
import openpyxl
import yaml
import io


YAML_HEAD_COMMENT = """# Classification item file {NAME_OF_FILE}
# {DESCRIPTION}
# https://circomod.eu/
"""


def parse_xlsx_and_save_as_yaml(xlsx_file):
    wb = openpyxl.load_workbook(xlsx_file)
    listSheet = [sheet for sheet in wb.sheetnames if re.match(r"CM_[^_]*_Definition", sheet)]
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

        stream = io.StringIO()
        yaml.dump(sheet_data, stream, sort_keys=False)
        stream_value = stream.getvalue()
        new_lines = []
        for line in stream_value.splitlines():
            if not line.startswith(" "):
                new_lines.append("")
            new_lines.append(line)

        with open('CIRCOMOD_Classification' + '-' + name + '-' + sheet_data["classification_info"][
            'classification_Name'] + '.yaml'
                , 'w') as file:
            file.writelines(YAML_HEAD_COMMENT.format(NAME_OF_FILE=sheet_data["classification_info"]
            ['classification_Name'], DESCRIPTION=sheet_data["classification_info"]['description']) + "\n".join(
                new_lines))


# Usage example
if __name__ == "__main__":
    XLSX_FILE = 'CIRCOMOD_Project_Wide_Classifications.xlsx'
    parse_xlsx_and_save_as_yaml(XLSX_FILE)
