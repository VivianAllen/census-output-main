#!/usr/bin/env python
# -*- coding: utf-8 -*-

import csv
import json
import re
import sys
import unicodedata

from openpyxl import load_workbook

CONFIG_WORKSHEET = "INDEX-filtered"
TOPIC_NAME_COLUMN = "Topic Area(s)"
VARIABLE_CODE_COLUMN = "2021 Mnemonic (variable)"
CLASS_TO_KEEP_COLUMN = "Classifications to keep"
DEFAULT_CLASS_COLUMN = "Default classification"
DOT_DENSITY_CLASS_COLUMN = "Dot density classification"
COMPARISON_2011_COLUMN = "2011 comparability?"

NON_ENTITIES = ("return to index", "does not apply")


# ==================================================== UTILITIES ===================================================== #


def cmp_strings(str1, str2):
    return str1.lower().strip() == str2.lower().strip()


def cmp_string_to_list(string, strList):
    return any(cmp_strings(string, str2) for str2 in strList)


def slugify(value, allow_unicode=False):
        """
        Convert to ASCII if 'allow_unicode' is False. Convert spaces or repeated
        dashes to single dashes. Remove characters that aren't alphanumerics,
        underscores, or hyphens. Convert to lowercase. Also strip leading and
        trailing whitespace, dashes, and underscores.

        * Copied from Django source *
        https://github.com/django/django/blob/f825536b5e09b3a047fec0c10aabd91bace0995c/django/utils/text.py#L400-L417
        """
        value = str(value)
        if allow_unicode:
            value = unicodedata.normalize("NFKC", value)
        else:
            value = (
                unicodedata.normalize("NFKD", value)
                .encode("ascii", "ignore")
                .decode("ascii")
            )
        value = re.sub(r"[^\w\s-]", "", value.lower())
        return re.sub(r"[-\s]+", "-", value).strip("-_")


def load_metadata():
    meta = {}
    for key, meta_file in {
        "topics": "Topic.csv", 
        "variables": "Variable.csv", 
        "classifications": "Classification.csv",
        "categories": "Category.csv"
    }.items():
        with open(meta_file, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            meta[key] = list(reader)
    return meta    


# ================================================= TOPIC PROCESSING ================================================= #


def get_topics(wb, metadata):
    config_rows = worksheet_to_row_dicts(wb[CONFIG_WORKSHEET])
    topics = []
    for cr in config_rows:
        name_or_mnemonic = cr[TOPIC_NAME_COLUMN].value
        if name_or_mnemonic is None:
            print(f"No topic name found for row {list(filter(lambda x: x, [c.value for c in cr.values()]))}, cannot process")
            continue
        topic_content = get_topic_content(name_or_mnemonic, metadata)
        # each topic row is also a variable row, so get those
        topic_row_variable = get_variable(wb, metadata, cr)
        seen_topic = next(filter(lambda t: topic_content["name"] == t["name"], topics), None)
        if seen_topic:
            seen_topic["variables"].append(topic_row_variable)
        else:
            topic_content["variables"].append(topic_row_variable)
            topics.append(topic_content)
    return topics


def worksheet_to_row_dicts(ws):
    rows = ws.rows
    header_row = next(rows)
    colnames = [cell.value for cell in header_row]
    return [dict(zip(colnames, row)) for row in rows if not all (c.value == None for c in row)]


def get_topic_content(name_or_mnemonic, metadata):
    topic_metadata = [
        m for m in metadata["topics"] if cmp_string_to_list(
            name_or_mnemonic, (m["Topic_Mnemonic"], m["Topic_Description"], m["Topic_Title"])
        )
    ]
    if topic_metadata:
        topic_name = topic_metadata[0]["Topic_Title"].strip()
        topic_desc = topic_metadata[0]["Topic_Description"].strip()
    else:
        topic_name = name_or_mnemonic
        topic_desc = "not found in topic metadata!"
        print(f"No metadata found for topic {name_or_mnemonic}")
    return {
        "name": topic_name,
        "slug": slugify(topic_name),
        "desc": topic_desc,
        "variables": []
    }

# =============================================== VARIABLE PROCESSING ================================================ #


def get_variable(wb, metadata, config_row):
    var_code_cell = config_row[VARIABLE_CODE_COLUMN]
    if var_code_cell.hyperlink is None:
        print(f"Ignoring variable {var_code_cell.value} as it does not link to a variable in spreadsheet.")
        return
    var_code_from_hyperlink = var_code_cell.hyperlink.location.split('!')[0].replace("'", "")
    variable_content = get_variable_content(var_code_from_hyperlink, metadata)
    variable_content["classifications"] = get_classifications(wb, metadata, variable_content, config_row)
    
    default_class_suffix = config_row[DEFAULT_CLASS_COLUMN].value.replace("(only one classification)", "").strip()
    default_class = class_code_from_suffix(variable_content["classifications"], default_class_suffix)
    if default_class:
        variable_content["default_classification"] = default_class[0]
    else:
        variable_content["default_classification"] = "not found in variables!"
        print(f"Default classification {default_class_suffix} could not be found in for variable {variable_content['code']}")
    
    dot_density_class_suffix = config_row[DOT_DENSITY_CLASS_COLUMN].value.strip()
    if dot_density_class_suffix.lower() == "no":
        variable_content["dot_density_classification"] = "false"
    else:
        dot_density_class = class_code_from_suffix(variable_content["classifications"], dot_density_class_suffix)
        if dot_density_class:
            variable_content["dot_density_classification"] = dot_density_class[0]
        else:
            variable_content["dot_density_classification"] = "not found in variables!"
            print(f"Dot density classification {dot_density_class_suffix} could not be found in for variable {variable_content['code']}")

    comp_2011 = config_row[COMPARISON_2011_COLUMN].value
    if comp_2011:
        variable_content["2011_comparison"] = comp_2011.replace("no", "false").replace("yes", "true")
    else:
        variable_content["2011_comparison"] = "false"

    return variable_content


def get_variable_content(code, metadata):
    var_metadata = [m for m in metadata["variables"] if cmp_strings(code, m["Variable_Mnemonic"])]
    if var_metadata:
        var_name = var_metadata[0]["Variable_Title"].strip()
        var_desc = var_metadata[0]["Variable_Description"].strip()
        var_units = var_metadata[0]["Statistical_Unit"].strip()
    else:
        var_name = code.replace("_", " ").strip().title()
        var_desc = "not found in variable metadata!"
        var_units = "not found in variable metadata!"
        print(f"No metadata found for variable {code}")
    return {
        "name": var_name,
        "code": code,
        "slug": slugify(var_name),
        "desc": var_desc,
        "units": var_units,
        "classifications": []
    }


def class_code_from_suffix(classifications, suffix):
    return [c["code"] for c in classifications if c["code"].lower().endswith(suffix.lower())]


# ============================================ CLASSIFICATION PROCESSING ============================================= #

def get_classifications(wb, metadata, variable, config_row):
    class_worksheet = wb[variable["code"]]
    col_headers = [cell.value for cell in next(class_worksheet.rows)]
    all_class_in_worksheet = [h for h in col_headers if isinstance(h, str) and h.lower() not in NON_ENTITIES]

    required_class_str = config_row[CLASS_TO_KEEP_COLUMN].value
    if required_class_str == "all":
        class_to_process = all_class_in_worksheet
    else:
        required_class = [x.strip() for x in required_class_str.split(",")]
        class_to_process = [c for c in all_class_in_worksheet if any(c.endswith(rc) for rc in required_class)]
    
    class_data_cols_index = [{"code": c, "cat_q_codes_col": i, "cat_name_col": i+1} for i, c in enumerate(col_headers) if c in class_to_process]
    classifications = []
    for c in class_data_cols_index:
        class_content = get_classification_content(c["code"], metadata)
        class_content["categories"] = get_categories(class_worksheet, c["cat_q_codes_col"], c["cat_name_col"], metadata, config_row)
        classifications.append(class_content)
    return classifications


def get_classification_content(code, metadata):
    code = code.replace(" ", "")
    class_metadata = [m for m in metadata["classifications"] if cmp_strings(code, m["Classification_Mnemonic"])]
    if class_metadata:
        class_desc = class_metadata[0]["External_Classification_Label_English"].strip()
    else:
        class_desc = "not found in classification metadata!"
        print(f"No metadata found for classification {code}")
    return {
        "code": code,
        "slug": slugify(code),
        "desc": class_desc,
        "categories": []
    }


# ================================================ CATEGORY PROCESSING =============================================== #


def get_categories(ws, cat_q_codes_col, cat_name_col, metadata, config_row):
    cat_q_codes = [norm_cat_q_codes(cell.value) for cell in list(ws.columns)[cat_q_codes_col] if is_bordered_cell(cell)]
    cat_names = [cell.value for cell in list(ws.columns)[cat_name_col] if is_bordered_cell(cell)]
    categories = []
    for q_codes, name  in zip(cat_q_codes, cat_names):
        if name.lower() not in NON_ENTITIES:
            categories.append(
                {
                    "name": name,
                    "slug": slugify(name),
                    "code": make_cat_code(q_codes, name)
                }
            )
    return categories


def norm_cat_q_codes(cat_q_code_str):
    cat_q_code_str = "".join(str(cat_q_code_str).split())
    for to_replace in ("–", ">"):
        cat_q_code_str = cat_q_code_str.replace(to_replace, "-")
    return cat_q_code_str


def is_bordered_cell(cell):
    return cell.border.left.style is not None and cell.border.right.style is not None


def make_cat_code(cat_q_codes, cat_name):
    return f"{slugify(cat_name)}={cat_q_codes}" 


# ======================================================= MAIN ======================================================= #

def main():
    workbook_filename = sys.argv[1]
    wb = load_workbook(workbook_filename)
    metadata = load_metadata()
    topics = get_topics(wb, metadata)
    with open(sys.argv[2], "w") as f:
        json.dump(topics, f, indent=4)


if __name__ == "__main__":
    main()
