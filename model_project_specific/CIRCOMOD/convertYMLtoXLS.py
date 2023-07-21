import re
import xlsxwriter
import yaml
import pandas as pd


def parse_yaml_and_save_as_xlsx(yaml_file):

    dfMetaTitle = pd.DataFrame(data={'Column': {'Metadata': 'Metadata'}})


    with open(yaml_file, 'r') as data:
        doc = yaml.load(data, Loader=yaml.FullLoader)

    dfCol = []
    dfColMeta = []
    dfValItem_Des = pd.DataFrame(data=range(1, len(doc['classification_items_description'])+1), columns=['id'])

    writer = pd.ExcelWriter("dict1.xlsx", engine="xlsxwriter")

    for i in doc.keys():
        if i == 'classification_info':
            dfCol = pd.DataFrame(data=doc[i].keys(), columns=['Column'])
            dfVal = pd.DataFrame(data=doc[i].values())
            dfCol['Value'] = dfVal

        if i == 'metadata':
            dfColMeta = pd.DataFrame(data=doc[i].keys(), columns=['Column'])
            dfValMeta = pd.DataFrame(data=doc[i].values())
            dfColMeta['Value'] = dfValMeta

    dfCol = pd.concat([dfCol, dfMetaTitle],axis=0)
    print(dfCol)
    dfClass_Def = pd.concat([dfCol, dfColMeta])
    dfClass_Def.to_excel(writer, sheet_name=doc["classification_info"]['id'] + '_Definition', index=False)

    for key in doc.keys():
        if 'items' in key:
            dfValItem_Des_temp = pd.DataFrame(data=doc[key].values(), columns=[key])
            dfValItem_Des = pd.concat([dfValItem_Des, dfValItem_Des_temp], axis=1)

    dfValItem_Des.to_excel(writer, sheet_name=doc["classification_info"]['id'] + '_Items', index=False)
    writer.close()


if __name__ == "__main__":
    YAML_FILE = 'CIRCOMOD_Classification-CM_1-material_processing_technologies.yaml'
    parse_yaml_and_save_as_xlsx(YAML_FILE)
