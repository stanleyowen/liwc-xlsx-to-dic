import pandas as pd
import sys


def build_category_tree_from_xlsx(df):
    # Row indices (0-based):
    # 0: 詞數 (word count)
    # 1: Parent
    # 2: 類別編號 (category id)
    # 3: 變項簡寫 (category short name)
    # 4: 變項名稱 (category description)
    # Column 0 contains row labels, actual data starts from column 1
    ids = df.iloc[2, 1:]
    names = df.iloc[3, 1:]
    descs = df.iloc[4, 1:]
    parents = df.iloc[1, 1:]
    # Build category dict
    cats = {}
    for idx, (cid, name, desc, parent) in enumerate(zip(ids, names, descs, parents)):
        if pd.isna(cid):
            continue
        # Convert to int then string to handle numeric ids properly
        try:
            cid = str(int(float(cid)))
        except:
            continue
        if pd.isna(name):
            continue
        # Handle parent id
        parent_id = None
        if not pd.isna(parent):
            try:
                parent_id = str(int(float(parent)))
                # If parent equals self, it's a root category
                if parent_id == cid:
                    parent_id = None
            except:
                pass
        cats[cid] = {
            'id': cid,
            'name': str(name),
            'desc': str(desc) if not pd.isna(desc) else '',
            'parent': parent_id,
            'children': []
        }
    # Build tree
    root_cats = []
    for cat in cats.values():
        parent_id = cat['parent']
        if parent_id and parent_id in cats:
            cats[parent_id]['children'].append(cat)
        else:
            root_cats.append(cat)
    return cats, root_cats


def write_category_tree_section(f, root_cats, indent=0):
    if indent == 0:
        f.write('%\n')
    for cat in sorted(root_cats, key=lambda x: int(x['id'])):
        f.write(f"{'    '*indent}{cat['id']}\t{cat['name']} ({cat['desc']})\n")
        if cat['children']:
            write_category_tree_section(f, cat['children'], indent+1)
    if indent == 0:
        f.write('%\n')


def extract_word_category_pairs(df):
    # Read all columns starting from column 1
    ids = df.iloc[2, 1:]
    word_category_pairs = []
    for idx, cid in enumerate(ids):
        if pd.isna(cid):
            continue
        try:
            cid = str(int(float(cid)))
        except:
            continue
        col = 1 + idx
        words = df.iloc[5:, col].dropna().astype(str).values.flatten().tolist()
        for word in words:
            word = word.strip()
            if word:
                word_category_pairs.append((word, cid))
    return word_category_pairs


def group_words_by_categories(word_category_pairs):
    from collections import defaultdict
    word_to_cats = defaultdict(list)
    for word, cat in word_category_pairs:
        if cat not in word_to_cats[word]:
            word_to_cats[word].append(cat)
    return word_to_cats


def xlsx_to_liwc_dic(xlsx_path, dic_path):
    df = pd.read_excel(xlsx_path, header=None)
    cats, root_cats = build_category_tree_from_xlsx(df)
    word_category_pairs = extract_word_category_pairs(df)
    word_to_cats = group_words_by_categories(word_category_pairs)

    with open(dic_path, 'w', encoding='utf-8') as f:
        write_category_tree_section(f, root_cats)
        for word, cats in sorted(word_to_cats.items()):
            cat_str = '\t'.join(str(c) for c in sorted(cats, key=int))
            f.write(f"{word}\t{cat_str}\n")


if __name__ == '__main__':
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print('Usage: python xlsx_to_liwc_dic.py input.xlsx [output.dic]')
        sys.exit(1)

    input_file = sys.argv[1]

    if len(sys.argv) == 3:
        output_file = sys.argv[2]
    else:
        # Generate output filename by replacing extension with .dic
        import os
        base_name = os.path.splitext(input_file)[0]
        output_file = base_name + '.dic'

    xlsx_to_liwc_dic(input_file, output_file)
