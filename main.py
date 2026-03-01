import os
from file_management import export_folder_to_xlsx_unified, add_languages, auto_format_excel, check_xlsx_columns, read_xlsx_raw, export_xlsx_to_txt

def txts_folder_to_xlsx(folder_path: str, output_pathfile: str, added_languages: list[str]) -> None:
    print(f"Reading TXT files from folder: {folder_path}")
    export_folder_to_xlsx_unified(folder_path, output_pathfile)
    print(f"Exported unified XLSX to: {output_pathfile}")

    print("Adding new languages:", ', '.join(added_languages), "(and Template)")
    added_languages.append("Template")
    add_languages(output_pathfile, added_languages)
    print("Languages added.")

    print("Auto-formatting Excel file")
    auto_format_excel(output_pathfile)
    print(f"Finished formatting: {output_pathfile}")

def xlsx_to_txts(xlsx_path: str, output_folder: str) -> None:
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

def main() -> None:
    run = True
    while run:
        print("\n")
        print("Type 1 to turn a folder with x.txt (folder with language txt files, like English.txt) into a spreadsheet")
        print("Type 2 to turn a x.xlsx into different language files (x.txt, like English.txt, among many options)")
        print("Type 0 to exit")
        action = int(input())

        if action == 0:
            run = False
        elif action == 1:
            print("What languages would you want to add? Write them separated by commas (no spaces, unless the language name has one) and capitalized")
            added_languages = input().split(",")
            print("Type the path to the TXT files")
            folder_path_txts = input()
            print("Type the output folder")
            output_pathfile_xlsx = os.path.join(input(), "export.xlsx")
            txts_folder_to_xlsx(folder_path_txts, output_pathfile_xlsx, added_languages)
        elif action == 2:
            print("Type the pathfile to the XLSX (NOTE: Include the name of the file and extension too, like ./import.xlsx)")
            xlsx_path = input()
            print("Type the output folder for the TXT")
            output_folder_txts = input()
            xlsx_to_txts(xlsx_path, output_folder_txts)
        else:
            print("Unknown command, try again")

if __name__ == "__main__":
    main()
