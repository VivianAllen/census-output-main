#!/usr/bin/env python

import csv
import json
import re
import sys
import unicodedata

def get_topics(variables):
    topics = set()
    for variable in variables:
        topics.add(variable["metadata"]["Topic Area(s)"])
    return sorted(topics)

def summarise_category(category):
    return {
        "name": category["name"],
        "slug": slugify(category["name"]),
        "codes": category["codes"],
        "desc": "TODO",
        "category_h_pt2": "TODO",
        "category_h_pt3": "TODO",
    }
def summarise_classification(classification, additional_metadata):
    additional_classification_metadata = next(
        filter(lambda x: x["Classification_Mnemonic"] == classification["classification_code"].lower(), additional_metadata["classifications"])
    )
    return {
        "code": classification["classification_code"],
        "desc": additional_classification_metadata["External_Classification_Label_English"],
        "categories": [summarise_category(c) for c in classification["categories"]],
        "default": classification["default"]
    }
def summarise_variable(variable, additional_metadata):
    metadata = variable["metadata"]
    additional_variable_metadata = next(
        filter(lambda x: x["Variable_Mnemonic"] == metadata["2021 Mnemonic (variable)"].lower(), additional_metadata["variables"])
    )
    return {
        "name": metadata["Description"],
        "slug": slugify(metadata["Description"]),
        "code": metadata["2021 Mnemonic (variable)"],
        "desc": additional_variable_metadata["Variable_Description"],
        "dataset": metadata["Dataset"], 
        "units": additional_variable_metadata["Statistical_Unit"],
        "classifications": [summarise_classification(c, additional_metadata) for c in variable["classifications"]]
    }
    
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

def get_vars_by_topic(topic, variables, additional_metadata):
    result = []
    for variable in variables:
        if variable["metadata"]["Topic Area(s)"] == topic:
            result.append(summarise_variable(variable, additional_metadata))
    return result

def load_additional_metadata():
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

def main():
    with open(sys.argv[1], "r") as f:
        variables = json.load(f)

    additional_metadata = load_additional_metadata()
    topics = get_topics(variables)
    result = [
        {
            "name": topic,
            "slug": slugify(topic),
            "desc": "TODO desc",
            "variables": get_vars_by_topic(topic, variables, additional_metadata)
        }
        for topic in topics
    ]
    with open(sys.argv[2], "w") as f:
        json.dump(result, f)

if __name__=="__main__":
    main()
