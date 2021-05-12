import re

def find_name(cell):
    m = re.compile(r"[a-z]+[ ][a-z]+|[a-z]+")
    cell = cell.strip().lower()
    result = m.match(cell)
    return result.group(0)

def cat_match(cell, category):
    cell = cell.lower().strip()
    cell = find_name(cell)
    for word in category:
        if re.match(word, cell):
            return True
        else:
            return False

def size_convert(size):
    translation = {'T':'Twin', 'TXL':'Twin XL', 'F':'Full', 'Q':'Queen', 'K':'King', 'STD':'Standard', 'E':'Euro'}
    if size in translation:
        return translation[size]
    else:
        return size

def size_label_find(cell):
    m = re.compile(r"[A-Z]+|[a-z]+|[A-Z][a-z]+")
    result = m.match(cell)
    if result.group(0).lower() == 'king':
        return 'K'
    return result.group(0)

def drop_unnamed(df):
    for col in df.columns:
        if col.startswith('Unn'):
            df.drop(columns=col,inplace=True)
    return df

def dim_find(cell):
    dims =  re.findall(r"[0-9]+", cell)
    dims = [x+' in' for x in dims]
    return dims

def colors_find(df, family_row, family_col):
    family_row += 1
    color = df.loc[family_row][family_col]
    try:
        if color:
            colors.append(color.upper())
            colors_find(family_row, family_col)
    except:
        pass
