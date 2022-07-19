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


def norm_string(string):
    return string.lower().strip()

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
        "classifications": 
        "Classification.csv", 
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
    return [dict(zip(colnames, row)) for row in rows]


def get_topic_content(name_or_mnemonic, metadata):
    topic_metadata = [
        m for m in metadata["topics"] if norm_string(name_or_mnemonic) in (
            norm_string(m["Topic_Mnemonic"]),
            norm_string(m["Topic_Description"]), 
            norm_string(m["Topic_Title"]),
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
    return variable_content


def get_variable_content(code, metadata):
    var_metadata = [m for m in metadata["variables"] if norm_string(code) == norm_string(m["Variable_Mnemonic"])]
    if var_metadata:
        var_name = var_metadata[0]["Variable_Title"].strip()
        var_desc = var_metadata[0]["Variable_Description"].strip()
    else:
        var_name = norm_string(code).replace("_", " ").title()
        var_desc = "not found in variable metadata!"
        print(f"No metadata found for variable {code}")
    return {
        "name": var_name,
        "code": code,
        "slug": slugify(var_name),
        "desc": var_desc,
        "classifications": []
    }


# ============================================ CLASSIFICATION PROCESSING ============================================= #

def get_classifications(wb, metadata, variable, config_row):
    class_worksheet = wb[variable["code"]]
    col_headers = [cell.value for cell in next(class_worksheet.rows)]
    all_class_in_worksheet = [h for h in col_headers if isinstance(h, str)]
    required_class_str = config_row[CLASS_TO_KEEP_COLUMN].value
    if required_class_str == "all":
        class_to_process = all_class_in_worksheet
    else:
        required_class = required_class_str.split(",")
        class_to_process = [c for c in all_class_in_worksheet if any(c.endswith(rc) for rc in required_class)]
    
    class_data_cols_index = [{"code": c, "cat_code_col": i, "cat_name_col": i+1} for i, c in enumerate(col_headers) if c in class_to_process]
    classifications = []
    for c in class_data_cols_index:
        class_content = get_classification_content(c["code"], metadata)
        classifications.append(class_content)
    return classifications

def get_classification_content(code, metadata):
    code = code.replace(" ", "")
    class_metadata = [m for m in metadata["classifications"] if norm_string(code) == norm_string(m["Classification_Mnemonic"])]
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


# ======================================================= MAIN ======================================================= #

def main():
    workbook_filename = sys.argv[1]
    wb = load_workbook(workbook_filename)
    metadata = load_metadata()
    topics = get_topics(wb, metadata)
    # variable_list = index_worksheet_to_variable_list(wb)
    with open(sys.argv[2], "w") as f:
        json.dump(topics, f, indent=4)


if __name__ == "__main__":
    main()
