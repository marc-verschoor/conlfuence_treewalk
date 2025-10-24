import json
import xlsxwriter
from pathlib import Path
import argparse

def generate_spreadsheet(input_file, config_object, docname=""):
    # Load JSON data from file
    try: 
        with open(input_file, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e: 
        print("Error reading file {input_file}, exception {e}. ")
        exit (1) 
    

    # Create a new Excel workbook
    if docname !="":
        output_file = docname
    else: 
        if 'xlsx_docname' in config_object:
            output_file = config_object['xlsx_docname']
        else: 
            output_file = Path(input_file).stem + "_.xlsx"
    workbook = xlsxwriter.Workbook(output_file)

    # Define formats
    wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top', 'font_name': 'Chillax'})
    alt_format_1 = workbook.add_format({'bg_color': '#F2F2F2', 'text_wrap': True, 'valign': 'top','font_name': 'Chillax'})
    alt_format_2 = workbook.add_format({'bg_color': '#FFFFFF', 'text_wrap': True, 'valign': 'top','font_name': 'Chillax'})
    alt_format_1_num = workbook.add_format({'num_format': '0', 'bg_color': '#F2F2F2', 'text_wrap': True, 'valign': 'top','font_name': 'Chillax'})
    alt_format_2_num = workbook.add_format({'num_format': '0','bg_color': '#FFFFFF', 'text_wrap': True, 'valign': 'top','font_name': 'Chillax'})
    header_format = workbook.add_format({'font_color': "#FFFFFF", 'bg_color': '#0000FF','bold': True, 'text_wrap': True, 'valign': 'top','font_name': 'Chillax', 'font_size': 12})
    def write_sheet_vertical(sheet_name, dict):
        # this will write a flat dict on the sheet
        row = 0
        worksheet = workbook.add_worksheet(sheet_name[:31])
        for key, value in dict.items():
            worksheet.write(row,0,key, header_format)
            if isinstance(value,int): 

                worksheet.write(row,1,value,alt_format_1_num)
            row+=1
        worksheet.set_column(0, 1, 80)
            

    def write_sheet(sheet_name, records,order):
        worksheet = workbook.add_worksheet(sheet_name[:31])  # Excel sheet name limit

        # Determine all unique keys
        all_keys = set()
        for record in records:
            all_keys.update(record.keys())
        all_keys = list(all_keys)

        # Move 'Notes', 'source_id', and 'doc_url' to the end
        for col in order:
            if col in all_keys:
                all_keys.remove(col)
                all_keys.append(col)

        # Write headers
        for col_num, key in enumerate(all_keys):
            worksheet.write(0, col_num, key, header_format)

        # Write data rows with alternating color based on source_title
        last_title = None
        use_alt_1 = True

        for row_num, record in enumerate(records, start=1):
            current_title = record.get("source_title", "")
            if current_title != last_title:
                use_alt_1 = not use_alt_1
                last_title = current_title
            row_format = alt_format_1 if use_alt_1 else alt_format_2

            for col_num, key in enumerate(all_keys):
                value = record.get(key, "")
                if isinstance(value, str):
                    value = value.replace("&lt;br&gt;", "\n").replace("<br>", "\n")
                if isinstance(value,int): 
                    if row_format == alt_format_1: 
                        worksheet.write(row_num, col_num, value, alt_format_1_num)
                    else: 
                        worksheet.write(row_num, col_num, value, alt_format_2_num)
                else:                       
                    worksheet.write(row_num, col_num, value, row_format)
                

        # Auto-adjust column widths
        for col_num, key in enumerate(all_keys):
            max_len = max(len(str(record.get(key, ""))) for record in records)
            worksheet.set_column(col_num, col_num, min(max_len + 2, 50))

    # Write each category to its own sheet
    categories=[]
    order_dict ={}
    for items in config_object["tables"]:
        categories.append(items['after_header_name'])
        if 'presentation_order' in items: 

            order_dict[items['after_header_name']] = items['presentation_order']
        else: 
            print(f"generate_speadsheet: warning: category item {items['after_header_name']} does not have a presentation order.")
            order_dict[items['after_header_name']] = [] 
        print(f"Categories: {categories}\nOrder_dict: {order_dict}")
    if 'other_content'in config_object:
        for items in config_object["other_content"] :
            categories.append(items['list_key_name'])
            if 'presentation_order' in items: 
                order_dict[items['list_key_name']] = items[ 'presentation_order']
            else:
                order_dict[items['list_key_name']] = []



    for category in categories: 
        write_sheet(category, data[category], order_dict[category])
    write_sheet_vertical("confluence_treewalk",data["confluence_requirements_treewalk"])
    # Close the workbook
    workbook.close()
    print(f"Spreadsheet created: {output_file}")

def main():
    parser = argparse.ArgumentParser(description="Confluence-treewalk spreadsheet generator. Use -c to identify configuration file.")
    
    # Mandatory string argument
    parser.add_argument(
        "-c", 
        "--config", 
        type=str, 
        required=True, 
        help="Configuration name (string, mandatory)."
    )

    # Optional numeric argument
    parser.add_argument(
        "-s", 
        "--startdoc", 
        type=int, 
        default=None, 
        help="Start document number for treewalk"
    )
    args = parser.parse_args()
    config_file = args.config

    # open config file 
    try: 
        with open(config_file, "rt") as file: 
            config_object = json.load(file)
    except Exception as e:
        print(f"Problem reading {config_file}: {e}")
        exit(1)

    if "START_DOC" in config_object:
        generate_spreadsheet(f"./confluence_treewalk_{config_object['START_DOC']}.json", config_object)
    else: 
        print("No START_DOC identfied in config file supplied.")

if __name__ == "__main__":
    main()
