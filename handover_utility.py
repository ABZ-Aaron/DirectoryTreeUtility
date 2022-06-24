import os
import pandas as pd
import numpy as np
import sys
from pathlib import Path
import xlsxwriter

def construct_excel_formulas(target_directory, dir_tree_df):
    """Find relative path of path in DataFrame, then construct Excel forumula that appends this to the
    path of the Excel workbook location. This is to ensure links don't break if drive letters change in SSD path.
    Also turn Document Number into Hyperlink.
    """
    dir_tree_df['Path'] = dir_tree_df['Path'].apply(lambda x: os.path.relpath(path = x, start = target_directory))
    try:
        dir_tree_df['Document No'] = ('=HYPERLINK(LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"' + dir_tree_df['Path'] + '","' + dir_tree_df['Document No'] + '")')
    except:
        pass
    dir_tree_df['Path'] = ('=HYPERLINK(LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"' + dir_tree_df['Path'] + '")')
    return dir_tree_df

def merge_with_metadata(dir_tree_df, merge_on = "File Name"):
    """Read in metadata from python script directory and left join"""
    try:
        meta_data_df = pd.read_excel(f'{PYTHON_PATH}/{METADATA}')
        

        res = [word for word in REQUIRED_METADATA_COLS if word not in meta_data_df.columns]
        if len(res) > 0:
            print(f"ERROR - Please check these column names are present in the metadata file: {res}")
            sys.exit()


        dir_tree_df = dir_tree_df.merge(meta_data_df, on = merge_on, how = 'left')
        return dir_tree_df
            
    except Exception as e:
        print(f"""ERROR - Please ensure file 'metadata.xls' exists in python script directory. 
        This should be downloaded from Aconex, and displays all metadata for files. 
        It should have several columns including 'File Name' and 'Document No'.
        Error is: {e}""")
        sys.exit(0)
    
def fill_doc_number(row):
    """If file doesn't have document number, set it as filename without suffix. 
    This is for non-aconex documents"""
    if str(row['Document No'])  == 'nan' and row['Category'] == 'File':
        return os.path.splitext(row['File Name'])[0]
    else:
        return row['Document No']

def get_target_directory():
    """Return command line argument which should be valid directory"""
    try:
        target_directory = sys.argv[1]
    except IndexError:
        print("ERROR - Did you pass the target path as a command line argument?")
        sys.exit(1)

    if target_directory.endswith("\\"):
        target_directory = target_directory[:-1]

    if not Path(target_directory).exists():
        print("ERROR - Target path not found. Exiting")
        sys.exit(1)

    return Path(target_directory + "\\")

def get_dir_tree(path):
    """Generate list of all file/folder/subfolder paths. Add details to 
    dictionary which is used throughout script"""
    path_details = []
    count = 0
    print(path)
    for root, dirs, files in os.walk(path, topdown = True):

        dirs[:] = [d for d in dirs if not d.startswith('$')]

        # Uncomment if you want to limit the results for testing
        #if count > 1000:
            #break

        folder = os.path.basename(os.path.normpath(root))
        relative_path = root.replace(path, '')
        relative_path_parts = Path(relative_path).parts

        print(f"Processing folder {folder}")

        # if these values are empty, we assume script as been run on a drive folder
        # e.g. C:\ or E:\. In which case we replace folder & relative path values with 
        # a hyperlink for current directory where excel resides.
        if relative_path == "" and folder == "":
            drive = '=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)'
            folder = relative_path = drive

        if relative_path == "":
            relative_path = folder

        dir_tree_dict = {"Path" : root, 
                    "Spaces" : len(relative_path_parts), 
                    "File Name" : folder, 
                    "Category" : 'Folder',
                    "Folder" : folder,
                    "RelativePath" : relative_path}
                        
        path_details.append(dir_tree_dict)

        for f in files:
            full_path = os.path.join(root, f)
            relative_path = full_path.replace(path, '')
            relative_path_parts = Path(relative_path).parts
            type = "File" if os.path.isfile(full_path) else "Folder"
            folder = os.path.basename(os.path.normpath(root))

            dir_tree_dict = {"Path" : full_path, 
                    "Spaces" : len(relative_path_parts), 
                    "File Name" : f, 
                    "Category" : type,
                    "Folder" : folder,
                    "RelativePath" : relative_path}

            path_details.append(dir_tree_dict)
            count += 1

    return path_details

def get_table_of_contents_list(path_details):
    """Get list which will be outputted as table of contents for all folders"""
    toc_list = []
    rel_list = []

    # We add leading spaces to folder name so it will visually represent 
    # table of contents better in Excel output.
    for dir_tree_dict in path_details:
        if dir_tree_dict["Category"] == "Folder":
            relative_path = dir_tree_dict['RelativePath']
            rel_list.append(relative_path)
            output = str(dir_tree_dict["Folder"]).rjust((dir_tree_dict["Spaces"] * 6) + len(dir_tree_dict["Folder"]), ".")
            toc_list.append(output)

    return rel_list, toc_list

def fill_in_file_type(row):
    """If File column exists, add in additional file extensions"""
    try:
        if (row['File'] != row['File']) and (row['Category'] == 'File'):
            extension = Path(row['File Name']).suffix.lstrip('.')
            row['File'] = extension
    except:
        pass
    return row

def save_to_excel(toc_df, data_df, target_directory):
    """Save dataframes to Excel using Xlsxwriter"""
    with pd.ExcelWriter(f"{PYTHON_PATH}/temp/temp.xlsx", engine = 'xlsxwriter') as writer:  

        # Write dataframes to Excel
        toc_df.to_excel(writer, sheet_name='Contents', index = False)
        data_df.to_excel(writer, sheet_name='Data', index = False)

        workbook = writer.book

        # Hide first col of contents page and set width of second col
        worksheet = writer.sheets['Contents']
        worksheet.set_column('A:A', None, None, {'hidden': 1})
        worksheet.set_column('B:B', 70, None)

        text = 'NOTE\n\n* Double click table of contents values to jump to relevant folder in Data tab\n\n' \
            '* In Data tab, click on value under "Document No" or "Path" to open folder/file\n\n' \
            '* Data tab can be filtered, but avoid sorting asc/desc as this will cause issues'
        worksheet.insert_textbox(1, 3, text, {'width': 300,'height': 200})
 

        # Colour rows where f column (Type) is equal to 'Folder'
        worksheet = writer.sheets['Data']
        folder_format = workbook.add_format({'bg_color': '#FFC7CE', 
                                            'font_color': '#9C0006', 
                                            'bottom' : 1, 
                                            'top' : 1, 
                                            'bold' : True})

        number_rows = len(data_df.index) + 1 
        col_letter = xlsxwriter.utility.xl_col_to_name(len(data_df.columns) - 1)
        worksheet.conditional_format(f"$A$2:${col_letter}${number_rows}",
                                      {"type": "formula",
                                       "criteria": '=INDIRECT("D"&ROW())="Folder"',
                                       "format": folder_format})

        file_format = workbook.add_format({'border' : 1})
        worksheet.conditional_format(f"$A$2:${col_letter}${number_rows}",
                                      {"type": "formula",
                                       "criteria": '=INDIRECT("D"&ROW())="File"',
                                       "format": file_format})

        # Add a header format.
        header_format = workbook.add_format({'bold': True,'bottom': 2,'bg_color': '#ADD8E6', 'left' : 1, 'right' : 1, 'align' : 'center'})

        # Write the column headers with the defined format.
        for col_num, value in enumerate(data_df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Freeze top row
        worksheet.freeze_panes(1, 0)

        # Hide first 4 cols, and set width of remaining
        worksheet.set_column('A:D', None, None, {'hidden': 1})
        worksheet.set_column('E:E', 75, None)
        worksheet.set_column('F:F', 35, None)
        worksheet.set_column('G:G', 35, None)
        worksheet.set_column('H:H', 10, None)
        worksheet.set_column('I:AZ', 25, None)
        
        """
        Xlswriter won't accept a XLSM extenstion, so we overwrite our temp file
        Xlswriter won't save this as it doesn't contain a macro, so we first extract VbaProject.bin macro 
        file from a real xlsm file using 'python vba_extract.py vbaProject.xlsm
        We then insert this. This VbaProject macro file also happens to contain the macro our workbook needs."""
        workbook.filename = target_directory + OUTPUT_FILE
        workbook.add_vba_project(f'{PYTHON_PATH}/vbaProject.bin')

def main():

    target_directory = get_target_directory()

    print("Getting Directory Tree...")
    path_details = get_dir_tree(target_directory)

    print("Initialising DataFrame...")
    dir_tree_df = pd.DataFrame(path_details)

    print("Merging Metadata...")
    dir_tree_df = merge_with_metadata(dir_tree_df, merge_on="File Name")

    print("Creating Table of Contents...")
    relative_path_list, table_of_contents_list = get_table_of_contents_list(path_details)
    table_of_contents_df = pd.DataFrame({'Relative path' : relative_path_list, 'Table of Contents' : table_of_contents_list})

    print("Additional Setup...")
    dir_tree_df.replace('', np.nan, inplace = True)
    dir_tree_df['Document No'] = dir_tree_df.apply(lambda row: fill_doc_number(row), axis=1)
    dir_tree_df = dir_tree_df.apply(fill_in_file_type, axis = 1)

    print("Generating hyperlinks...")
    dir_tree_df = construct_excel_formulas(target_directory, dir_tree_df)
    
    print("Reordering & Renaming columns...")
    dir_tree_df.rename(columns = {'File Name':'Basename', 'File':'Extension'}, inplace = True)
    cols_to_move = ['RelativePath', 'Spaces', 'Folder', 'Category', 'Path', 'Basename', 'Document No', 'Extension']
    dir_tree_df = dir_tree_df[ cols_to_move + [ col for col in dir_tree_df.columns if col not in cols_to_move ] ]

    print("Saving to Excel...")
    save_to_excel(table_of_contents_df, dir_tree_df, target_directory)

    print(f"SUCCESS - Complete. Output file stored under {target_directory}. Please don't move from this location")

if __name__ == "__main__":
    
    # Global Variables
    PYTHON_PATH = os.path.dirname(os.path.realpath(__file__))
    OUTPUT_FILE = 'Handover Index.xlsm'
    REQUIRED_METADATA_COLS = ['Document No', 'File Name', 'File']
    METADATA = "metadata.xls"

    # Main Script
    main()
