import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

def check_extension(filepath: str, extensions: str | set[str] | list[str] | tuple[str]) -> bool:
    """
    Checks if the file extension of a path is correct.

    Args:
        filepath (str): path of the file 
        extensions (str | set[str] | list[str] | tuple[str]): extension or multiple ones to check

    Returns:
        bool: True if the path ends with one of the extensions or the single extension
    """

    filepath = filepath.lower()

    if isinstance(extensions, str):
        extensions = (extensions,)

    return filepath.endswith(tuple(extensions))

def get_filename(filepath: str) -> str:
    """
    Gets the filename without the path or extension.

    Args:
        filepath (str): path of the file

    Returns:
        str: the filename
    """
    filename = filepath.replace("\\", "/").split("/")[-1]
    return filename.rsplit(".", 1)[0]

def read_txt(filepath: str) -> list[str]:
    """
    Reads a .txt file and gets its contents as a list, each item being a different line.

    Args:
        filepath (str): path of the TXT file 

    Returns:
        list[str]: each line of the .txt as a list

    Raises:
        ValueError: if the file extension isn't .txt
    """

    if not check_extension(filepath, ".txt"):
        raise ValueError(f"File extension of {filepath} must be .txt")
    
    with open(filepath, 'r', encoding="utf-8") as file:
        return file.readlines()

def interpret_txt_data(lines: list[str]) -> list[dict[str, str | None | list[dict[str,str]]]]:
    """
    Parses a .txt into a list of items. Each item starts with a key (could be None).
    Each line under the key becomes a dict with text, comment, and extra.
    
    Args:
        lines (list[str]): each line of the .txt as a list
    
    Returns:
        list[dict]: each dict has:
            "key": str | None
            "content": list of dicts {text, comment, extra}
    """
    result = []
    current_item = None  # dict with key and content list

    for line in lines:
        line = line.strip()

        # Split comment
        parts = line.split('#', 1)
        data = parts[0].strip()
        comment = parts[1].strip() if len(parts) > 1 else ""

        # Detect key
        if '=' in data:
            key, value = data.split('=', 1)
            key = key.strip()
            value = value.strip()

            # Save previous item if exists
            if current_item is not None:
                result.append(current_item)

            current_item = {"key": key, "content": []}

            # Add value if exists
            if value or comment or data:
                current_item["content"].append({
                    "text": value,
                    "comment": comment,
                    "extra": ""
                })
        else:
            # Line is either extra, comment only, or just text
            text = ""
            extra = ""
            if data.startswith('[') and data.endswith(']'):
                extra = data
            else:
                text = data

            # If no current item yet, key is None
            if current_item is None:
                current_item = {"key": None, "content": []}

            current_item["content"].append({
                "text": text,
                "comment": comment,
                "extra": extra
            })

    if current_item is not None:
        result.append(current_item)

    return result

def interpreted_to_dataframe(elements: list[dict[str, str | None | list[dict[str,str]]]], language: str = "English") -> tuple[pd.DataFrame, list[tuple[int, int]]]:
    """
    Converts interpreted data into a pandas DataFrame for spreadsheet export.
    
    Columns: Key, Extra, Comments, Contents
    Only the first row for each key contains the key and extra (for merging later).
    Returns a list of tuples (start_row, end_row) for each key block to merge later.
    
    Args:
        elements (list[dict]): output of interpret_txt_data
        language (str): language label for comments and contents columns
    
    Returns:
        tuple:
            - pd.DataFrame: formatted DataFrame
            - list of tuples: row ranges to merge the 'key' column later
    """
    rows = []
    merge_ranges = []
    current_row = 0

    for item in elements:
        key = item.get("key")
        contents_list = item.get("content", [])
        
        if not contents_list:
            # Ensure at least one empty row
            contents_list = [{"text": "", "comment": "", "extra": ""}]
        
        start_row = current_row
        for i, line in enumerate(contents_list):
            row = [
                key if i == 0 else None,  # Only first row gets the key
                line.get("extra", ""),
                line.get("comment", ""),
                line.get("text", "")
            ]
            rows.append(row)
            current_row += 1
        end_row = current_row - 1
        merge_ranges.append((start_row, end_row))  # tuple to merge "key" later

    df = pd.DataFrame(
        rows,
        columns=["Key", "Extra", f"Comments", f"Contents"]
    )
    return df, merge_ranges

def unify_keys(elements_by_language: dict[str, list[dict]]) -> dict[str, list[dict]]:
    """
    Ensures all languages have the same set of keys.
    Missing keys in any language are added with empty content.
    
    Args:
        elements_by_language (dict): {language_name: elements} as output of interpret_txt_data
    
    Returns:
        dict: updated {language_name: elements} with all keys present
    """
    # Collect all keys from all languages
    all_keys = set()
    for elements in elements_by_language.values():
        for item in elements:
            key = item.get("key")
            if key is not None:
                all_keys.add(key)

    # Add missing keys to languages
    unified = {}
    for lang, elements in elements_by_language.items():
        existing_keys = {item.get("key") for item in elements if item.get("key") is not None}
        new_elements = elements.copy()

        for key in all_keys - existing_keys:
            # Add a new item with the missing key and empty content
            new_elements.append({
                "key": key,
                "content": [{"text": "", "comment": "", "extra": ""}]
            })
        unified[lang] = new_elements

    return unified

def export_folder_to_xlsx_unified(folder_path: str, output_pathfile: str) -> None:
    """
    Reads all .txt files in a folder, parses them as languages, unifies all keys,
    and exports a multi sheet XLSX with merged key rows.
    
    Each .txt file is treated as a language, using the filename (without extension)
    as the sheet name.
    
    Args:
        folder_path (str): folder containing .txt files
        output_pathfile (str): path to save the XLSX workbook
    """
    # Read all languages first
    elements_by_language = {}
    for filename in os.listdir(folder_path):
        if not check_extension(filename, ".txt"):
            continue
        filepath = os.path.join(folder_path, filename)
        language_name = get_filename(filename)

        lines = read_txt(filepath)
        elements = interpret_txt_data(lines)
        elements_by_language[language_name] = elements

    # Unify keys across all languages
    elements_by_language = unify_keys(elements_by_language)

    # Convert to DataFrames with merge ranges
    sheets_data = {}
    for lang, elements in elements_by_language.items():
        df, merge_ranges = interpreted_to_dataframe(elements, language=lang)
        sheets_data[lang] = (df, merge_ranges)

    # Write to Excel
    with pd.ExcelWriter(output_pathfile, engine="openpyxl") as writer:
        for sheet_name, (df, _) in sheets_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Merge keys in Excel
    wb = load_workbook(output_pathfile)
    for sheet_name, (_, merge_ranges) in sheets_data.items():
        ws = wb[sheet_name]
        key_col = 1  # column A
        for start_row, end_row in merge_ranges:
            if start_row != end_row:
                ws.merge_cells(
                    start_row=start_row + 2,  # +2 for header row
                    start_column=key_col,
                    end_row=end_row + 2,
                    end_column=key_col
                )
    wb.save(output_pathfile)

def add_languages(xlsx_path: str, new_languages: list[str], base_language: str = None) -> None:
    """
    Adds new languages to an existing XLSX workbook.
    Copies keys from the base language but keeps the original content/comments/extra
    in additional columns for reference.
    
    Args:
        xlsx_path (str): path to the XLSX workbook
        new_languages (list[str]): list of new languages to add as sheets
        base_language (str, optional): sheet to copy keys/content from, defaults to first sheet
    """
    wb = load_workbook(xlsx_path)
    
    # Determine base sheet
    if base_language is None:
        base_sheet = wb.sheetnames[0]
    else:
        if base_language not in wb.sheetnames:
            raise ValueError(f"Base language sheet '{base_language}' not found.")
        base_sheet = base_language
    
    ws_base = wb[base_sheet]

    # Read rows from base sheet
    base_rows = []
    current_key = ""
    new_key = False

    for row in ws_base.iter_rows(min_row=2, values_only=True):
        key, extra, comment, content = row[:4]  # original 4 columns

        if key not in (None, ""):
            current_key = key
            new_key = True
        elif new_key:
            new_key = False

        base_rows.append({
            "key": current_key if new_key else "",
            "extra": extra if extra else "",
            "comment": comment if comment else "",
            "content": content if content else ""
        })
    
    # Add new sheets
    for lang in new_languages:
        if lang in wb.sheetnames:
            continue
        ws = wb.create_sheet(title=lang)
        
        # Header row: keep original 4 columns + 4 new ones
        ws.append([
            "Key", "Extra", f"Comments", f"Contents",
            "", "Translation notes", "Original Content", "Original Extra", "Original Comments"
        ])
        
        # Write rows
        for row_data in base_rows:
            ws.append([
                row_data["key"],
                "",
                "",
                "",
                "",
                "",
                row_data["content"],
                row_data["extra"],
                row_data["comment"]
            ])
    
        # Merge Key column automatically
        start_row = None

        for row in range(2, ws.max_row + 1):  # start after header
            cell_value = ws.cell(row=row, column=1).value

            if cell_value:  # new key starts
                if start_row and row - 1 > start_row:
                    ws.merge_cells(start_row=start_row, start_column=1,
                                end_row=row - 1, end_column=1)
                start_row = row

        # Merge last block
        if start_row and ws.max_row > start_row:
            ws.merge_cells(start_row=start_row, start_column=1,
                        end_row=ws.max_row, end_column=1)

    wb.save(xlsx_path)

def auto_format_excel(xlsx_path: str) -> None:
    """
    Adjusts column widths, enables text wrap for all sheets, 
    and formats the top row (headers) with bold font and freeze.
    
    Args:
        xlsx_path (str): path to the XLSX workbook
    """
    wb = load_workbook(xlsx_path)
    for ws in wb.worksheets:
        # Freeze top row
        ws.freeze_panes = ws['A2']

        for col_cells in ws.columns:
            max_length = 0
            column = col_cells[0].column_letter
            for i, cell in enumerate(col_cells):
                if cell.value:
                    # Treat as string and split lines
                    lines = str(cell.value).split("\n")
                    cell_max = max(len(line) for line in lines)
                    if cell_max > max_length:
                        max_length = cell_max
                
                # Wrap text
                cell.alignment = Alignment(wrap_text=True, vertical='top')
                
                # Bold header
                if i == 0:
                    cell.font = Font(bold=True)
            
            # Set column width with padding
            ws.column_dimensions[column].width = max_length + 5
    
    wb.save(xlsx_path)

def check_xlsx_columns(xlsx_path: str, required_columns=None) -> dict[str, bool]:
    """
    Checks if each sheet in the XLSX has the required columns in the first row.

    Args:
        xlsx_path (str): path to the XLSX file
        required_columns (list, optional): list of columns to check. Defaults to 
                                           ["Key", "Extra", "Comments", "Contents"]

    Returns:
        dict: {sheet_name: True if all columns exist, False otherwise}
    """
    if required_columns is None:
        required_columns = ["Key", "Extra", "Comments", "Contents"]

    wb = load_workbook(xlsx_path, data_only=True)
    results = {}

    for ws in wb.worksheets:
        headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]
        exists = all(col in headers for col in required_columns)
        results[ws.title] = exists

    return results

def read_xlsx_raw(xlsx_path: str) -> dict[str, list[dict[str, str]]]:
    """
    Reads an XLSX workbook and extracts only the columns: Key, Extra, Comments, Contents.
    Ignores formatting and other columns. Empty cells are returned as empty strings.

    Args:
        xlsx_path (str): path to the XLSX file

    Returns:
        dict: {sheet_name: list of rows}, where each row is a dict with keys:
              "Key", "Extra", "Comments", "Contents"
    """
    wb = load_workbook(xlsx_path, data_only=True)
    result = {}

    for ws in wb.worksheets:
        # Find header row (assume first row)
        headers = [cell.value for cell in ws[1]]
        # Map column indexes for required fields
        col_map = {}
        for idx, h in enumerate(headers):
            if h in ("Key", "Extra", "Comments", "Contents"):
                col_map[h] = idx

        rows_data = []
        # Read all rows starting from row 2
        for row in ws.iter_rows(min_row=2, values_only=True):
            row_dict = {}
            for key in ("Key", "Extra", "Comments", "Contents"):
                col_idx = col_map.get(key)
                row_dict[key] = str(row[col_idx]) if col_idx is not None and row[col_idx] is not None else ""
            rows_data.append(row_dict)

        result[ws.title] = rows_data

    return result

def export_xlsx_to_txt(xlsx_data: dict[str, list[dict[str, str]]], output_folder: str) -> None:
    """
    Exports XLSX data (from read_xlsx_raw) to separate TXT files per sheet.
    
    Rules:
    1. If Key exists, write 'KEY= ' first.
    2. Only one of 'Contents' or 'Extra' can exist per line. If both, raises an Error.
    3. Comment, if exists, goes after a '# '.
    4. Keep empty lines.
    
    Args:
        xlsx_data (dict): {sheet_name: list of row dicts with Key, Extra, Comments, Contents}
        output_folder (str): folder to save the TXT files
    """
    import os
    os.makedirs(output_folder, exist_ok=True)

    for sheet_name, rows in xlsx_data.items():
        if sheet_name == "Template":
            continue

        txt_lines = []

        for row in rows:
            key = row.get("Key", "").strip()
            extra = row.get("Extra", "").strip()
            content = row.get("Contents", "").strip()
            comment = row.get("Comments", "").strip()

            line_parts = []

            if key:
                line_parts.append(f"{key}=")

            if extra and content:
                raise ValueError(f"Row in sheet '{sheet_name}' has both 'Extra' and 'Contents'.")
            elif content:
                line_parts.append(content)
            elif extra:
                line_parts.append(extra)

            if comment:
                line_parts.append(f"# {comment}")

            line = " ".join(line_parts)

            if not line_parts:
                line = ""

            txt_lines.append(line)

        output_path = os.path.join(output_folder, f"{sheet_name}.txt")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(txt_lines) + "\n")
