#!/usr/bin/env python
# -*- coding: utf-8 -*-

import csv
from dataclasses import dataclass
import json
import re
import sys
import unicodedata

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


# ====================================================== CONFIG ====================================================== #


# csv files from sharepoint with descriptions, units etc for all topics, variables, clasifications & categories
METADATA_FILES = {
    "topics": "Topic.csv",
    "variables": "Variable.csv",
    "classifications": "Classification.csv",
    "categories": "Category.csv"
}
# csv file encoding. NB - might want to check this if you get weird characters in the metadata output
METADATA_FILE_ENCODING = "utf-8-sig'"

# The name of the sheet in the Output_Category_Mapping excel workbook with the variables to process and the additional
# config required to process them into atlas content files.
CONFIG_WORKSHEET = "INDEX-filtered"

# The column name (assumed to be first row) in the index page which specifies the topic each variable belongs to
TOPIC_NAME_COLUMN = "Topic Area(s)"

# The column name (assumed to be first row) in the index page which contains within-sheet hyperlinks to variable sheets
# to be included (e.g. Accommodation_Type, which links to the ACCOMODATION_TYPE worksheet). NB - all values in here that
# are NOT hyperlinks will not currently be processed, as they are assumed to refer to variables that are not
# defined in the Output_Category_Mapping excel workbook.
VAR_HYPERLINK_COLUMN = "2021 Mnemonic (variable)"

# The column name (assumed to be first row) that lists the classifications from each variable that are to be included.
# This can be either a single value (e.g. 2A), a comma-seperated list (e.g 2A, 4A, 5A) or 'all' (all defined
# classifications will be included.)
CLASSIFICATIONS_TO_INCLUDE_COLUMN = "Classifications to keep"

# The column name (assumed to be first row) that defines the default classification to be used for each variable.
CHOROLPLETH_DEFAULT_CLASS_COLUMN = "Default classification"

# The column name (assumed to be first row) that defines the classification for each variable that can be represented
# as a dot density map
DOT_DENSITY_DEFAULT_CLASS_COLUMN = "Dot density classification"

# The column name (assumed to be first row) that flags if this variable has comparison data from the previus 2011 census
COMPARISON_2011_COLUMN = "2011 comparability?"

# Values that are often found in the same place as data but are not data, and so shouldn't be included.
NOT_DATA = ("return to index", "does not apply", "no code required")


# ==================================================== CLASSES ======================================================= #


@dataclass
class ConfigRow:
    """All needed config values from a row from the config page of the input workbook (e.g INDEX-fitered)"""
    topic: str
    variable: Cell
    classifications: str
    chorolpleth_default_classification: str
    dot_density_default_classification: str
        

@dataclass
class CensusCategory:
    """A category as found in content.json."""
    name: str
    slug: str
    code: str

    def to_jsonable(self):
        """Category in json-friendly form."""
        return {
            "name": self.name,
            "slug": self.slug,
            "code": self.code
        }


@dataclass
class CensusClassification:
    """A classification as found in content.json."""
    code: str
    slug: str
    desc: str
    chorolpleth_default: bool = False
    dot_density_default: bool = False
    categories: list[CensusCategory] = []

    def to_jsonable(self):
        """Classification json-friendly form w. optional properties."""
        output_params = {
            "code": self.code,
            "slug": self.slug,
            "desc": self.desc
        }
        if self.chorolpleth_default:
            output_params["chorolpleth_default"] = self.chorolpleth_default
        if self.dot_density_default:
            output_params["dot_density_default"] = self.dot_density_default
        output_params["categories"] = [c.to_jsonable() for c in self.categories]
        return output_params


@dataclass
class CensusVariable:
    """A variable as found in content.json."""
    name: str
    code: str
    slug: str
    desc: str
    units: str
    classifications: list[CensusClassification] = []

    def to_jsonable(self):
        """Variable in json-friendly form."""
        return {
            "name": self.name,
            "code": self.code,
            "slug": self.slug,
            "desc": self.desc,
            "units": self.units,
            "classifications": [c.to_jsonable() for c in self.classifications]
        }


@dataclass
class CensusTopic:
    """A topic as found in content.json."""
    name: str
    slug: str
    desc: str
    variables: list[CensusVariable] = []

    def to_jsonable(self):
        """Topic in json-friendly form."""
        return {
            "name": self.name,
            "slug": self.slug,
            "desc": self.desc,
            "variables": [v.to_jsonable() for v in self.variables]
        }

@dataclass
class AllTopics:
    """A list of topics as found in content.json"""
    topics: list[CensusTopic]

    def to_jsonable(self):
        """Topic list in json-friendly form, with topics sorted alphabetically.."""
        return [t.to_jsonable() for t in sorted(self.topics, key = lambda x: x.name)]


# ==================================================== UTILITIES ===================================================== #


def load_config_rows_from_config_sheet(wb: Workbook) -> list[ConfigRow]:
    """
    To load_config_rows_from_config_sheet, first use the columns headers in the first row to convert all subsequent rows
    to dictionaries, then use row_dicts to create to list of ConfigRows. NB blank rows will be ignored.
    """
    rows = wb[CONFIG_WORKSHEET].rows
    header_row = next(rows)
    colnames = [cell.value for cell in header_row]
    row_dicts = [dict(zip(colnames, row)) for row in rows if not all(c.value == None for c in row)]
    config_rows = []
    for row_dict in row_dicts:
        config_rows.append(
            ConfigRow(
                topic = row_dict[TOPIC_NAME_COLUMN].value,
                variable = row_dict[VAR_HYPERLINK_COLUMN],
                classifications = row_dict[CLASSIFICATIONS_TO_INCLUDE_COLUMN],
                chorolpleth_default_classification = row_dict[CHOROLPLETH_DEFAULT_CLASS_COLUMN],
                dot_density_default_classification = row_dict[DOT_DENSITY_DEFAULT_CLASS_COLUMN]
            )
        )
    return config_rows
    

def load_topic_defin


def cmp_strings(string1: str, string2: str) -> bool:
    """Return True if normalised strings are equal."""
    return string1.lower().strip() == string2.lower().strip()


def cmp_string_to_list(string1: str, strList: str) -> bool:
    """Return True if normalised string1 is equal to any normalised string in strList."""
    return any(cmp_strings(string1, string2) for string2 in strList)


def slugify(value: str, allow_unicode=False) -> str:
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


def load_metadata() -> dict:
    """
    Load all csv files in the METADATA_FILES constant to a dictionary keyed for each file, containing lists of 
    dictionaries, each dictionary containing a single row's values keyed to the csv column headers.
    """
    meta = {}
    for key, meta_file in METADATA_FILES.items():
        with open(meta_file, "r", encoding=METADATA_FILE_ENCODING) as f:
            reader = csv.DictReader(f)
            meta[key] = list(reader)
    return meta


def worksheet_to_row_dicts(ws: Worksheet) -> list[dict]:
    


# ================================================= TOPIC PROCESSING ================================================= #


def get_topics(wb: Workbook, config_rows: list[ConfigRow], metadata: dict) -> AllTopics:
    """
    To parse the topics from the wb workbook, first loop through the ConfigRows (which define which topics, variables
    and classifications are to be processed) and first get_topic_metadata for each topic name found in the TOPIC_NAME_COLUMN (rows 
    with no value in this column will be ignored), then get_variable for each row.

    NB - topics and variables are defined in the same rows on the config worksheet, so the same topic will be processed
    multiple times. Only the first defintion of the topic will be saved.
    """
    topics = []
    for cr in config_rows:
        # the config sheet seems to refer to topics by either name, mnemonic or title...
        name_mnemonic_or_title = cr[TOPIC_NAME_COLUMN].value
        if name_mnemonic_or_title is None:
            print(
                f"No topic name found for row {list(filter(lambda x: x, [c.value for c in cr.values()]))}, cannot process")
            continue
        topic_metadata = get_topic_metadata(name_mnemonic_or_title, metadata)

        # each topic row is also a variable row, so get variable info here
        topic_variable = get_variable(wb, cr, metadata)

        # add this topic if we've not seen it before...
        topic_names = [t.name for t in topics]
        if topic_metadata["name"] not in topic_names:
            topics.append(
                CensusTopic(
                    name=topic_metadata["name"],
                    slug=slugify(topic_metadata["name"]),
                    desc=topic_metadata["desc"],
                    variables=[topic_variable]
                )
            )
        # ...otherwise if we have seen it before, add its variable to the existing topic
        else:
            topic = next(
                filter(lambda t: topic_metadata["name"] == t.name, topics))
            topic.variables.append(topic_variable)

    return AllTopics(topics = topics)


def get_topic_metadata(name_mnemonic_or_title: str, metadata: dict) -> dict:
    """
    To get metadata for a topic, search the 'topics' value of the metadata dict (a list of dictionaries) for 
    one in which the name, mnemonic or title matches the name_mnemonic_or_title arg, then retrieve the topic name
    and description from the metadata. If not found, substitute with warning string and print warning.
    """
    topic_metadata = [
        m for m in metadata["topics"] if cmp_string_to_list(
            name_mnemonic_or_title, (m["Topic_Mnemonic"],
                                     m["Topic_Description"], m["Topic_Title"])
        )
    ]
    if len(topic_metadata) > 0:
        topic_name = topic_metadata[0]["Topic_Title"].strip()
        topic_desc = topic_metadata[0]["Topic_Description"].strip()
    else:
        topic_name = name_mnemonic_or_title
        topic_desc = "not found in topic metadata!"
        print(f"No metadata found for topic {name_mnemonic_or_title}")

    return {
        "name": topic_name,
        "desc": topic_desc,
    }


# =============================================== VARIABLE PROCESSING ================================================ #


def get_variable(wb: Workbook, config_row: dict, metadata: dict) -> CensusVariable:
    """
    To parse a variable from the wb workbook, first get_variable_code from the config_row (return if code 
    could not be found), then get_variable_metadata for the variables code, then finally get_classifications from the 
    worksheet for the variable.
    """
    var_code = get_variable_code(config_row)
    if var_code is None:
        return

    var_metadata = get_variable_metadata(var_code, metadata)

    return CensusVariable(
        name=var_metadata["name"],
        code=var_code, slug=slugify(var_code),
        desc=var_metadata["desc"],
        units=var_metadata["units"],
        classifications=get_required_classifications(
            wb[var_code], config_row, metadata)
    )


def get_variable_code(config_row: dict) -> str or None:
    """
    To get the variable code from the config_row, read the value of the VAR_HYPERLINK_COLUMN cell. If this is not a 
    hyperlink, the config_row does not refer to the a variable that is defined in the current workbook, so return None.
    If it is a hyperlink, return the name of the worksheet it points to.
    """
    var_code_cell = config_row[VAR_HYPERLINK_COLUMN]
    if var_code_cell.hyperlink is None:
        print(
            f"Ignoring variable {var_code_cell.value} as it does not link to a variable in spreadsheet.")
        return None
    return var_code_cell.hyperlink.location.split('!')[0].replace("'", "")


def get_variable_metadata(code: str, metadata: dict):
    """
    To get metadata for a variable, search the 'variables' value of the metadata dict (a list of dictionaries) for 
    one in which the mnemonic matches the code arg, then retrieve the variable name, desc, and units from the metadata.
    If not found, substitute with warning string and print warning.
    """
    var_metadata = [m for m in metadata["variables"] if cmp_strings(code, m["Variable_Mnemonic"])]
    if len(var_metadata) > 0:
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
        "desc": var_desc,
        "units": var_units,
    }


# ============================================ CLASSIFICATION PROCESSING ============================================= #


def get_required_classifications(cls_ws: Worksheet, config_row: dict, metadata: dict):
    """
    To parse required classifications from the cls_ws worksheet, first get_classification_column_indices, then 
    filter_required_classifications to only those required in the config_row, then for each required classification,
    then convert the cls_ws.columns generator into an indexable list and get_classification_metadata, 
    get_classification_visualisation_flags, and then get_categories from those columns containing classification info.
    """
    cls_col_indices = get_classification_column_indices(cls_ws)
    required_cls_col_indices = filter_required_classifications(
        cls_col_indices, config_row)
    cls_ws_cols = [col for col in cls_ws.columns]
    classifications = []
    for c in required_cls_col_indices:
        cls_metadata = get_classification_metadata(c["cls_code"], metadata)
        cls_flags = get_classification_visualisation_flags(c["cls_code"], config_row)
        classifications.append(
            CensusClassification(
                code = cls_metadata["code"],
                slug = slugify(cls_metadata["code"]),
                desc = cls_metadata["desc"],
                chorolpleth_default = cls_flags["chorolpleth_default"],
                dot_density_default = cls_flags["dot_density_default"],
                categories = get_categories(cls_ws_cols[c["cat_codes_col"]], cls_ws_cols[c["cat_name_col"]])
            )
        )
    return classifications


def get_classification_column_indices(cls_ws: Worksheet) -> list[dict]:
    """
    To get classification column indices, first get the classification codes by extracting all non-empty strings from 
    the first row of the worksheet and filtering out those listed in the NOT_DATA constant, then get the column index
    of each classification code (the category codes are assumed to also in this column) and the column index immediately
    to its right (the category names are assumed to be in this column).
    """
    col_headers = [cell.value for cell in next(cls_ws.rows)]
    cls_headers = [h for h in col_headers if isinstance(
        h, str) and not cmp_string_to_list(h, NOT_DATA)]
    return [
        {"cls_code": c, "cat_codes_col": i, "cat_name_col": i+1} for i, c in enumerate(col_headers) if c in cls_headers
    ]


def filter_required_classifications(cls_column_indices: list[dict], config_row: dict) -> list[dict]:
    """
    To filter required classifications, first extract and parse the value of the CLASSIFICATIONS_TO_INCLUDE_COLUMN in 
    the config_row dict (which may be 'all' or a comma-seperated list of classification suffixes) and filter 
    the cls_column_indices dict to only those referenced in this value.
    """
    required_cls_str = config_row[CLASSIFICATIONS_TO_INCLUDE_COLUMN].value
    if required_cls_str == "all":
        return cls_column_indices
    else:
        required_cls = [x.strip() for x in required_cls_str.split(",")]
        return [c for c in cls_column_indices if any(c["cls_code"].endswith(rc) for rc in required_cls)]


def get_classification_metadata(code: str, metadata: dict) -> dict:
    """
    To get metadata for a classification, search the 'classifications' value of the metadata dict (a list of 
    dictionaries) for one in which the mnemonic matches the code arg, then retrieve the variable name, desc, and units 
    from the metadata. If not found, substitute with warning string and print warning.
    """
    code = code.replace(" ", "")
    class_metadata = [m for m in metadata["classifications"]
                      if cmp_strings(code, m["Classification_Mnemonic"])]
    if class_metadata:
        class_desc = class_metadata[0]["External_Classification_Label_English"].strip(
        )
    else:
        class_desc = "not found in classification metadata!"
        print(f"No metadata found for classification {code}")
    return {
        "code": code,
        "desc": class_desc,
    }


def get_classification_visualisation_flags(code: str, config_row: dict) -> dict:
    default_class_suffix = config_row[CHOROLPLETH_DEFAULT_CLASS_COLUMN].value.replace("(only one classification)", "").strip()
    dot_density_class_suffix = config_row[DOT_DENSITY_DEFAULT_CLASS_COLUMN].value.strip()
    return {
        "chorolpleth_default": code.endswith(default_class_suffix),
        "dot_density_default": dot_density_class_suffix.lower() != "no" and code.endswith(dot_density_class_suffix)
    }


# ================================================ CATEGORY PROCESSING =============================================== #


def get_categories(cat_codes_col: tuple[Cell], cat_name_col: tuple[Cell]) -> list[CensusCategory]:
    """
    To parse categories from cat_codes_col and cat_name_col worksheet columns, first get_category_cell_indices each 
    (throws exception if these do not match), then create_category_code from both the codes and the 
    name for each category. NB - category names found in the NOT_DATA constant will be ignored.
    """
    cat_code_cell_indices = get_category_cell_indices(cat_codes_col)
    cat_name_cell_indices = get_category_cell_indices(cat_name_col)

    if cat_code_cell_indices != cat_name_cell_indices:
        error_str = (
            f"Fatal error processing classification {cat_codes_col[0].value} - code and name columns don't match up!"
            " Differences seen in cells "
            f"{[x + 1 for x in set(cat_code_cell_indices).difference(set(cat_name_cell_indices))]}."
        )
        raise Exception(error_str)
    
    categories = []
    for i in cat_code_cell_indices:
        if not cmp_string_to_list(cat_name_col[i].value, NOT_DATA):
            categories.append(
                CensusCategory(
                    name = cat_name_col[i].value,
                    slug = slugify(cat_name_col[i].value,),
                    code = make_cat_code(cat_name_col[i].value, cat_codes_col[i].value)
                )
            )
    return categories


def get_category_cell_indices(col: tuple[Cell]) -> list[int]:
    """
    To get category cell indices from a col, find indices of all populated cells that have left or right borders defined 
    (this seems to be a consistent style used in the Output_Category_Mapping excel workbook). NB - enumerate from second
    row (making sure to adjust the index) and ignore the first row as this should be the category code, regardless of 
    formatting.
    """
    category_cell_indices = []
    for i, c in enumerate(col[1:], 1):
        if c.value != None and (c.border.left.style is not None or c.border.right.style is not None):
            category_cell_indices.append(i)
    return category_cell_indices

def make_cat_code(cat_name, cat_codes):
    """
    To make a unqiue category_code from the cat_name and cat_codes of a category, first normalise the cat_codes 
    statement to have no whitespace and consist only of numbers, commas and '-' (with no leading zeroes) then combine with 
    the slugified cat_name, seperated by an =.
    """
    cat_codes = "".join(str(cat_codes).split())
    for to_replace in ("–", ">"):
        cat_codes = cat_codes.replace(to_replace, "-")
    if cat_codes[0] == "0":
        cat_codes = cat_codes[1:]
    return f"{slugify(cat_name)}={cat_codes}"


# ======================================================= MAIN ======================================================= #


def main():
    workbook_filename = sys.argv[1]
    output_filename = sys.argv[2]
    wb = load_workbook(workbook_filename)
    config_rows = load_config_rows_from_config_sheet(wb)
    # metadata = load_metadata()
    # topics = get_topics(wb, config_rows, metadata)
    # with open(output_filename, "w") as f:
    #     json.dump(topics.to_jsonable(), f, indent=4)


if __name__ == "__main__":
    main()
