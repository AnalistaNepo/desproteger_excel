import zipfile
import shutil
import os
import re

def unprotect_excel(file_path):
    """
    Remove sheet and workbook protection from an Excel file (.xlsx or .xlsm).
    Creates a new file with '_unprotected' suffix.
    """
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' not found.")
        return

    file_name, file_ext = os.path.splitext(file_path)
    output_path = f"{file_name}_unprotected{file_ext}"

    # Create a temporary directory to extract files
    temp_dir = "temp_excel_unprotect"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    try:
        print(f"Processing '{file_path}'...")
        
        # Extract all files from the Excel zip archive
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Process worksheets to remove sheet protection
        worksheets_dir = os.path.join(temp_dir, "xl", "worksheets")
        if os.path.exists(worksheets_dir):
            for sheet_file in os.listdir(worksheets_dir):
                if sheet_file.endswith(".xml"):
                    full_path = os.path.join(worksheets_dir, sheet_file)
                    with open(full_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # Regex to remove <sheetProtection ... />
                    # It handles self-closing tags and attributes within the tag
                    new_content = re.sub(r'<sheetProtection[^>]*/>', '', content)
                    
                    if content != new_content:
                        print(f"  - Removed protection from {sheet_file}")
                        with open(full_path, 'w', encoding='utf-8') as f:
                            f.write(new_content)

        # Process workbook.xml to remove workbook protection
        workbook_path = os.path.join(temp_dir, "xl", "workbook.xml")
        if os.path.exists(workbook_path):
            with open(workbook_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Regex to remove <workbookProtection ... />
            new_content = re.sub(r'<workbookProtection[^>]*/>', '', content)
            # Also try without self-closing if it has content (less common for protection but possible)
            new_content = re.sub(r'<workbookProtection[^>]*>.*?</workbookProtection>', '', new_content, flags=re.DOTALL)

            if content != new_content:
                print("  - Removed workbook protection from workbook.xml")
                with open(workbook_path, 'w', encoding='utf-8') as f:
                    f.write(new_content)

        # Re-zip the files into the new Excel file
        print(f"Creating unprotected file: '{output_path}'")
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path_on_disk = os.path.join(root, file)
                    # Archive name should be relative to temp_dir
                    archive_name = os.path.relpath(file_path_on_disk, temp_dir)
                    zip_out.write(file_path_on_disk, archive_name)
        
        print("Done! You can now open the _unprotected file.")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Cleanup temp directory
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

if __name__ == "__main__":
    # Target file
    target_file = r"c:\Users\Administrador\Desktop\Desproteger Excel\Termômetro Saúde Operacional T2 - Janeiro.xlsx"
    unprotect_excel(target_file)
