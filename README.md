# LIWC XLSX to DIC Converter

A Python tool to convert LIWC (Linguistic Inquiry and Word Count) category data from Excel (.xlsx) format into the standard LIWC dictionary (.dic) format.

## Overview

This converter reads a specially formatted Excel file containing LIWC categories and word lists, then outputs a properly formatted .dic file that includes:

- A hierarchical category tree with parent-child relationships
- Word-category mappings where each word is associated with its relevant category IDs

## Requirements

- Python 3.6 or higher
- pandas
- openpyxl

## Installation

1. **Clone or download this repository**

2. **Create a virtual environment (recommended):**

   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On macOS/Linux
   # or
   venv\Scripts\activate  # On Windows
   ```

3. **Install dependencies:**
   ```bash
   pip install pandas openpyxl
   ```

## Usage

### Basic Usage

```bash
python xlsx_to_liwc_dic.py input.xlsx [output.dic]
```

### Arguments

- `input.xlsx` (required): Path to the input Excel file
- `output.dic` (optional): Path to the output .dic file
  - If not provided, the output file will have the same name as the input file with a `.dic` extension

### Examples

**Example 1: Auto-generate output filename**

```bash
python xlsx_to_liwc_dic.py cliwc2015_v1.6.xlsx
```

Creates: `cliwc2015_v1.6.dic`

**Example 2: Specify output filename**

```bash
python xlsx_to_liwc_dic.py cliwc2015_v1.6.xlsx my_dictionary.dic
```

Creates: `my_dictionary.dic`

**Example 3: Using relative paths**

```bash
python xlsx_to_liwc_dic.py ./data/input.xlsx ./output/result.dic
```

## Input File Format

The Excel file must follow this specific structure:

### Row Structure (0-indexed)

- **Row 0**: Word count (Number of words for each category)
- **Row 1**: Parent (Parent category ID)
- **Row 2**: Category ID (Unique numeric identifier for the category)
- **Row 3**: Category Code (Category short name/code)
- **Row 4**: Category Name (Category description/full name)
- **Row 5+**: Word list (one word per row)

### Column Structure

- **Column 0**: Row labels (Word count, Parent, Category ID, etc.)
- **Column 1+**: Each column represents a category with its metadata and word list

### Example Structure

```
Word Count  | 56      | 56       | 45      | 11  | 6   | 9   | 10    | 4    |
Parent      | 1       | 1        | 2       | 3   | 3   | 3   | 3     | 3    |
Category ID | 1       | 2        | 3       | 4   | 5   | 6   | 7     | 8    |
Cat. Code   | func    | pronoun  | ppron   | i   | we  | you | shehe | they |
Cat. Name   | Function| Pronouns | Personal| I   | We  | You | ShHe  | They |
------------|---------|----------|---------|-----|-----|-----|-------|------|
            | my      | my       | my      | my  | we  | you | he    | they |
            | mine    | mine     | mine    | I   | us  | your| she   | them |
            | ...     | ...      | ...     | ... | ... | ... | ...   | ...  |
```

### Category Hierarchy Rules

- **Root categories**: When a category's Parent ID equals its own Category ID (e.g., Category 1 with Parent 1), it's treated as a root category
- **Child categories**: When a category's Parent ID differs from its Category ID, it becomes a child of that parent
- **Indentation**: The output automatically indents child categories based on their depth in the hierarchy

## Output File Format

The generated .dic file follows the LIWC dictionary format:

### Structure

```
%
[Category Tree Section]
%
[Word-Category Mappings]
```

### Category Tree Section

Between the two `%` markers, categories are listed hierarchically:

```
%
1	func (Function Words)
    2	pronoun (Pronouns)
        3	ppron (Personal Pronouns)
            4	i (First Person Singular)
            5	we (First Person Plural)
            6	you (Second Person)
            7	shehe (Third Person Singular)
            8	they (Third Person Plural)
%
```

- Format: `[indent][id]TAB[name] ([description])`
- Indentation: 4 spaces per level of hierarchy

### Word-Category Mappings Section

After the second `%`, each line contains a word and its associated category IDs:

```
I	1	2	3	4
we	1	2	3	5
you	1	2	3	6
he	1	2	3	7
```

- Format: `[word]TAB[cat_id_1]TAB[cat_id_2]TAB...`
- Words are sorted alphabetically
- Category IDs are sorted numerically
- A word inherits all parent category IDs

## How It Works

1. **Parse Excel**: Reads the Excel file and extracts category metadata (ID, name, description, parent)
2. **Build Category Tree**: Constructs a hierarchical tree structure based on parent-child relationships
3. **Extract Words**: Collects all words from each category column
4. **Map Words to Categories**: Associates each word with its category and all ancestor categories
5. **Generate Output**: Writes the category tree and word mappings to the `.dic` file

## License

See LICENSE file for details.
