#!/usr/bin/env python
# -*- coding: utf-8 -*-

import csv
import json
import re
import sys
import unicodedata

from openpyxl import load_workbook




INDEX_WORKSHEET = "INDEX-filtered"
TOPIC_NAME_COLUMN = "Topic Area(s)"
INDEX_WORKSHEET_VARIABLE_CODE_COLUMN = "2021 Mnemonic (variable)"

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


def index_worksheet_to_row_dicts(index_worksheet):
    rows = index_worksheet.rows
    header_row = next(rows)
    colnames = [cell.value for cell in header_row]
    return [dict(zip(colnames, row)) for row in rows]


def get_seen_topic(topic_name, topics):
    return next(filter(lambda t: topic_name == t["name"], topics), None)


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
    

def get_topics(index_rows, metadata):
    topics = []
    for r in index_rows:
        name_or_mnemonic = r[TOPIC_NAME_COLUMN].value
        if name_or_mnemonic is None:
            print(f"No topic name found for row {list(filter(lambda x: x, [c.value for c in r.values()]))}, cannot process")
            continue
        topic_content = get_topic_content(name_or_mnemonic, metadata)
        # each topic row is also a variable row, so get those
        topic_row_variable = get_variable(metadata, r)
        seen_topic = get_seen_topic(topic_content["name"], topics)
        if seen_topic:
            seen_topic["variables"].append(topic_row_variable)
        else:
            topic_content["variables"].append(topic_row_variable)
            topics.append(topic_content)

    return topics


def get_variable_metadata(code, metadata):
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
        "slug": slugify(var_name),
        "desc": var_desc,
    }

def variable_code_from_hyperlink_cell(hyperlink_cell):
    return hyperlink_cell.hyperlink.location.split('!')[0].upper().replace("'", "")


def get_variable(metadata, row):
    var_code_cell = row[INDEX_WORKSHEET_VARIABLE_CODE_COLUMN]
    if var_code_cell.hyperlink is None:
        print(f"Ignoring variable {var_code_cell.value} as it does not link to a variable in spreadsheet.")
        return
    var_code = variable_code_from_hyperlink_cell(var_code_cell)
    variable_metadata = get_variable_metadata(var_code, metadata)
    return variable_metadata

def main():
    workbook_filename = sys.argv[1]
    wb = load_workbook(workbook_filename)
    metadata = load_metadata()
    index_rows = index_worksheet_to_row_dicts(wb[INDEX_WORKSHEET])
    topics = get_topics(index_rows, metadata)
    # variable_list = index_worksheet_to_variable_list(wb)
    with open(sys.argv[2], "w") as f:
        json.dump(topics, f, indent=4)


if __name__ == "__main__":
    main()
