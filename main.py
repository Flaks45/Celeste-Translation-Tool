import os
from file_management import export_folder_to_xlsx_unified, add_languages, auto_format_excel, check_xlsx_columns, read_xlsx_raw, export_xlsx_to_txt

def txts_folder_to_xlsx():
    folder_path = "./samples/test/"
    output_pathfile = "./samples/test/export.xlsx"

    print(f"Reading TXT files from folder: {folder_path}")
    export_folder_to_xlsx_unified(folder_path, output_pathfile)
    print(f"Exported unified XLSX to: {output_pathfile}")

    print("Adding new languages: Spanish (and Template)")
    add_languages(output_pathfile, ["Spanish", "Template"])
    print("Languages added.")

    print("Auto-formatting Excel file")
    auto_format_excel(output_pathfile)
    print(f"Finished formatting: {output_pathfile}")

def xlsx_to_txts():
    xlsx_path = "./samples/test/import.xlsx"
    output_folder = "./samples/test/results/"

    check_result = check_xlsx_columns(xlsx_path)
    all_ok = True
    for sheet, ok in check_result.items():
        print(f"{sheet}: {'OK' if ok else 'MISSING COLUMNS'}")
        if not ok:
            all_ok = False

    if not all_ok:
        print("One or more sheets are missing required columns. Aborting export.")
        return
        
    xlsx_data = read_xlsx_raw(xlsx_path)
    export_xlsx_to_txt(xlsx_data, output_folder)
    print("\nGenerated TXT files:")
    for sheet_name in xlsx_data.keys():
        if sheet_name == "Template":
            continue

        txt_path = os.path.join(output_folder, f"{sheet_name}.txt")
        print(txt_path)

if __name__ == "__main__":
    xlsx_to_txts()
    input("Press Enter to exit.")
