import os
import pandas as pd
import csv
from tcs_functions import *

#enter product categories as key and sheet mispellings as values
categories = {'Duvet Covers':['duvet cover', 'duvet covers'], 'Fitted Sheets':['fitted sheet', 'fitted sheets'],\
                'Flat Sheets':['flat sheet', 'flat sheets'], 'Shams':['sham', 'shams'], 'Pillowcases':['pillowcase', \
                'pillowcases'], 'Men\'s Long Robe':['men\'s long robe']}

#columns to be filled on final output
columns = ['Model Number', 'Product Category', 'Collection Name', 'Size', 'Length', 'Width', 'Color', 'Material', \
           'Certifications', 'Pattern Type', 'Theme', "Pocket Depth", 'Fiber Content', 'Machine Wash', 'Country of Origin', \
          'Attributes', 'Closure Type', 'Fill Type', 'Carton Length', 'Carton Width', 'Carton Height', 'Carton Weight',\
          'UPC', 'GTIN', 'Package Weight', 'Package Length', 'Package Width', 'Package Height']
columns_fill = {}
for x in columns:
    columns_fill[x] = None

created=[]
main_dict = {}
for file in os.scandir():
    if file.path.endswith('.xlsm'):
        df = pd.read_excel(file.path, 'Product Copy Sheet', skiprows=[0,1,2])
        df.columns = ["0", "1", "2", "3", "4", "5", "6"]

        try:
            df_setup = pd.read_excel(file.path, 'SETUP SHEET', skiprows=\
                                 [0,1,2,3,4,5,6,7,8,9,10])
            df_setup.columns = df_setup.iloc[0]
            df_setup = df_setup[1:]
            cols = [x for x in range(len(df_setup.columns))]
            df_setup.columns = cols
            df_setup.rename(columns={1:'Product', 3:'SKU', 4:'Size', 12:'Carton Length', 13:'Carton Width',\
                14: 'Carton Height', 15: 'Carton Weight', 46:'UPC', 47: 'GTIN', 49:'Package Weight', 50:'Package Length', \
                51:'Package Width',52:'Package Height'},inplace = True)
            df_setup.dropna(subset=['Product'], inplace=True)
            df_setup['SKU'] = df_setup['SKU'].astype(str)
            df_setup['SKU'] = df_setup['SKU'].astype(str)
            df_setup['Size'] = df_setup['Size'].astype(str)
        except:
            print('Problem with or missing setup sheet for {}'.format(file.path))

        collection_name = df.loc[df['0'] == "PRODUCT/COLLECTION NAME:"]['1'].values[0]
        fiber = df.loc[df['0'] ==  "FIBER CONTENT - TOTAL: (write in)"]['1'].values[0]
        fill = df.loc[df['0'] == "FILL TYPE: (drop down)"]['1'].values[0]
        pocket_depth = df.loc[df['0'] == "FITTED SHEET POCKET DEPTH: (write in)"]['1'].values[0]
        closure_type =  df.loc[df['0'] == "CLOSURE TYPE: (drop down)"]['1'].values[0]
        attributes = df.loc[df['0'] == "ATTRIBUTES: (Select ALL that apply)"]['1'].values[0]
        pattern_type =  df.loc[df['0'] == "PATTERN TYPE: (drop down)"]['1'].values[0]
        theme = df.loc[df['0'] == "THEME: (Select ALL that apply)"]['1'].values[0]
        certifications =  df.loc[df['0'] == "CERTIFICATIONS: (Select ALL that apply)"]['1'].values[0]
        country = df.loc[df['0'] == "COO: (drop down)"]['1'].values[0]
        material = df.loc[df['0'] == "MATERIAL: (drop down)"]['1'].values[0]

        for row in df.iterrows():
            items = [item for item in enumerate(row[1])]
            for num, item in items:
                if "FAMILY COLOR" == item:
                    family_col = num
                    family_row = row[0]
                    break

        colors = []
        def colors_find(family_row, family_col):
            family_row += 1
            color = df.loc[family_row][family_col]
            try:
                if color:
                    colors.append(color.upper())
                    colors_find(family_row, family_col)
            except:
                pass

        colors_find(family_row, family_col)

        for cat, typo in categories.items():
            cat = cat.lower()
            cat_dict = {}
            cat_list = []
            for row in df.iterrows():
                # 0 is row number, 1 is tuple position
                if type(row[1][0]) == str and row[1][0] != 'FITTED SHEET POCKET DEPTH: (write in)':
                    if cat_match(row[1][0], typo):
                        model_number = row[1][1]
                        size_label = size_label_find(row[1][2])
                        model_base = model_number + '-' + size_label
                        length, width = dim_find(row[1][2])
                        for color in colors:
                            full_model_number = model_base + '-' + color
                            cat_dict[full_model_number] = columns_fill
                            cat_dict[full_model_number]['Model Number'] = full_model_number
                            cat_dict[full_model_number]['Material'] = material
                            cat_dict[full_model_number]['Color'] = color
                            cat_dict[full_model_number]['Size'] = size_convert(size_label)
                            cat_dict[full_model_number]['Length'] = length
                            cat_dict[full_model_number]['Width'] = width
                            cat_dict[full_model_number]['Collection Name'] = collection_name
                            cat_dict[full_model_number]['Fiber Content'] = fiber
                            cat_dict[full_model_number]['Product Category'] = cat.title()
                            cat_dict[full_model_number]['Fill Type'] = fill
                            if cat == 'fitted sheets':
                                cat_dict[full_model_number]['Pocket Depth'] = pocket_depth
                            if cat == 'duvet covers' or cat == 'shams':
                                cat_dict[full_model_number]['Closure Type'] = closure_type
                            cat_dict[full_model_number]['Attributes'] = attributes
                            cat_dict[full_model_number]['Pattern Type'] = pattern_type
                            cat_dict[full_model_number]['Theme'] = theme
                            cat_dict[full_model_number]['Certifications'] = certifications
                            cat_dict[full_model_number]['Country of Origin'] = country
                            try:
                                df_setup
                                for setup_row in df_setup.iterrows():
                                    if setup_row[1]['SKU']+'-'+size_label_find(setup_row[1]['Size']) == model_base:
                                        cat_dict[full_model_number]['GTIN'] = str(int(setup_row[1]['GTIN']))
                                        cat_dict[full_model_number]['UPC']= str(int(setup_row[1]['UPC']))
                                        cat_dict[full_model_number]['Carton Length'] = str(setup_row[1]['Carton Length'])+' in'
                                        cat_dict[full_model_number]['Carton Width'] = str(setup_row[1]['Carton Width'])+' in'
                                        cat_dict[full_model_number]['Carton Height'] = str(setup_row[1]['Carton Height'])+' in'
                                        cat_dict[full_model_number]['Carton Weight'] = str(setup_row[1]['Carton Weight'])+' lb'
                                        cat_dict[full_model_number]['Package Weight'] = str(setup_row[1]['Package Weight'])+' lb'
                                        cat_dict[full_model_number]['Package Length'] = str(setup_row[1]['Package Length'])+' in'
                                        cat_dict[full_model_number]['Package Width'] = str(setup_row[1]['Package Width'])+' in'
                                        cat_dict[full_model_number]['Package Height'] = str(setup_row[1]['Package Height'])+' in'
                            except: pass

                            cat_list.append(cat_dict[full_model_number].copy())
                            for x in columns:
                                try:del cat_dict[full_model_number][x]
                                except:pass
            main_dict[cat] = cat_list
        empty_dict = [key for key, values in main_dict.items() if len(values) < 1]
        for key in empty_dict:
            del main_dict[key]

        for cat, specs in main_dict.items():
            try:
                temp_df = pd.read_excel('{}.xlsx'.format(cat.title()))
                final=pd.DataFrame(data=specs)
                final=pd.concat([final, temp_df])
                final=drop_unnamed(final)
                final.to_excel('{}.xlsx'.format(cat.title()))
            except:
                final=pd.DataFrame(data=specs)
                final=drop_unnamed(final)
                final.to_excel('{}.xlsx'.format(cat.title()))
                created.append(cat.title())

print('Created or updated sheets for {}.'.format(str(set(created))[1:-1]))
