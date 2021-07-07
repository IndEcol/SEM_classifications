# -*- coding: utf-8 -*-
"""
Created on Wed Jul  7 11:36:53 2021

@author: spauliuk
"""
# yaml classification parsing test

import yaml
P0 = 'C:\\Users\\spauliuk.AD\\FILES\\ARBEIT\\PROJECTS\\Database\\SEM_classifications\\resources_materials\\SEM_materials_Level1.yaml'
P1= 'C:\\Users\\spauliuk.AD\\FILES\\ARBEIT\\PROJECTS\\Database\\SEM_classifications\\models_projects\\ODYM_RECC\\SEM_materials_Level2_RECC_v2.5.yaml'

with open(P1, 'r') as stream:
    try:
        ClassFileContent = yaml.safe_load(stream)
    except yaml.YAMLError as exc:
        print(exc)
        
ClassFileContent['classification_items']
ClassFileContent['classification_info']
ClassFileContent['Metadata']        

# Convert info to IEDC upload info.