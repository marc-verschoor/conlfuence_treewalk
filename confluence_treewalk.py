# pip install browser_cookie3 requests
from datetime import datetime
import requests
import re
import json
import urllib3
import copy
from bs4 import BeautifulSoup
import argparse
import generate_spreadsheet
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
# this data is now in json config file.
#BASE_URL = "https://leaseplan-digital.atlassian.net/wiki/"
#USER_BASE_URL = "https://leaseplan-digital.atlassian.net/wiki/spaces/DCE/pages/"
#START_DOC = 4715577937

def increase_page_count(treewalk_dict):
    if not 'pagecount' in treewalk_dict['confluence_requirements_treewalk']:
        treewalk_dict['confluence_requirements_treewalk']['pagecount'] = 0
    treewalk_dict['confluence_requirements_treewalk']['pagecount'] +=1

def increase_element_count(treewalk_dict, element_type, by_num):
    if not f"element_count_{element_type}" in treewalk_dict['confluence_requirements_treewalk']:
        treewalk_dict['confluence_requirements_treewalk'][f"element_count_{element_type}"] = 0
    treewalk_dict['confluence_requirements_treewalk'][f"element_count_{element_type}"] +=by_num

def increase_attachment_count(treewalk_dict):
    if not 'attachment_count' in treewalk_dict['confluence_requirements_treewalk']:
        treewalk_dict['confluence_requirements_treewalk']['attachment_count'] = 0
    treewalk_dict['confluence_requirements_treewalk']['attachment_count'] +=1

def increase_unparsed_page_count(treewalk_dict):
    if not 'unparsed_page_count' in treewalk_dict['confluence_requirements_treewalk']:
        treewalk_dict['confluence_requirements_treewalk']['unparsed_page_count'] = 0
    treewalk_dict['confluence_requirements_treewalk']['unparsed_page_count'] +=1
    


def extract_cell_text(cell):
    # If the cell has no nested elements, return its plain text
    if not list(cell.children):
        return cell.get_text(strip=True)
    
    text_parts = []

    for elem in cell.descendants:
        if elem.name == "li":
            text_parts.append(f"- {elem.get_text(strip=True)}")
        elif elem.name == "p":
            # Avoid duplicating list text
            if not elem.find_parent("li"):
                text_parts.append(elem.get_text(strip=True))

    # Fallback: if nothing matched (like a plain number)
    if not text_parts:
        return cell.get_text(strip=True)

    return "<br>".join(text_parts)

def get_data_from_table(after_header_name, header_entries, soup, base_element, exclude_list = ["e.g."]):
    # check if the header is in the document
    print(f"== Getting data from table {after_header_name}: {header_entries}")
    header = soup.find(lambda tag: tag.name in ["h1","h2","h3"] and re.search(rf"{after_header_name}", tag.get_text(), re.I))
    
    if header: 
        next_table = header.find_next("table")
        # check if the table found has the right entries
        rows = next_table.find_all("tr")
        # format check
        element_list = []
        cells = [cell.get_text(strip=True) for cell in rows[0].find_all(["th", "td"])]
        headers = copy.deepcopy(cells)
        print(headers)
        elements = 0 
        if len(cells)>=len(header_entries): 
            for entry in header_entries: 
                if not entry in cells:
                    print(f"entry {entry} missing from table header {cells}, should be {header_entries} for {after_header_name}") 
                    return []
            for row in rows[1:]: 
                cells = [cell.get_text(strip=True) for cell in row.find_all(["th", "td"])]
                if cells[0] in exclude_list and not cells[0] in header_entries:
                    print(f"Cells[0] of row {cells} is in exclude list. value {cells[0]}")
                    continue
                else: 
                    
                    element = copy.deepcopy(base_element)
                    dont_include = False
                    for entry in header_entries: 
                        
                        if entry == "": 
                        # substituating an empty header for ID
                            element["id"] = cells[headers.index(entry)]
                        else: 
                            for exclude_item in exclude_list:
                                if cells[headers.index(entry)].startswith(exclude_item):
                                    dont_include = True
                                    continue
                            element[entry] = cells[headers.index(entry)]
                    if dont_include == False: 
                        element_list.append(element)
            return element_list 
                        
        else:
            print(f"not enough entries in table header {cells}, should be {header_entries} for {after_header_name}")
            return []
    else:
        print(f"Header {after_header_name} not found")


def retrieve_page_details_2(session, base_url, user_base_url, page_id, table_list, treewalk_dict): 
    response = session.get(f"{base_url}/rest/api/content/{page_id}?expand=body.storage", verify=False)
    data = response.json()
    if response.status_code != 200:
        print(f"Error retrieving page with id {page_id} - status code = {response.status_code}")
        return -1 
    data = response.json() 
    html = data["body"]["storage"]["value"]
    text = BeautifulSoup(html,"html.parser").get_text(separator="\n", strip=True)
    soup = BeautifulSoup(html, "html.parser")
    # this validates page structure - for now we are keeping this. 
    tables = soup.find_all("table")
    # Check if we are dealing with the right template. The first table MUST contain the following keywords:
    keywords = ['Target release',  'Epic', 'Document Status', 'Document Owner', 'Tech lead', 'Technical writers', 'QA']
    all_found = True
    increase_page_count(treewalk_dict)
    if len(tables)>0:
        for keyword in keywords: 
            if tables[0].find(keyword) == 0:
                all_found = False
                break
        if all_found: 
            for table in table_list: 
                base_element = { }
                base_element['source_title'] = data['title']
                base_element['source_id'] = page_id
                base_element['doc_url'] = f"{user_base_url}/{page_id}"
                elements_list = get_data_from_table(table['after_header_name'], table['table_entries'],soup,base_element)
                if elements_list: 
                    increase_element_count(treewalk_dict,table['after_header_name'],len(elements_list))
                if not table['after_header_name'] in treewalk_dict:
                    treewalk_dict[table['after_header_name']] = []
                treewalk_dict[table['after_header_name']].extend(elements_list)
        if "_expandable" in data: 
                print (f"retrieving URL for page property 'children' {data['_expandable']['children']}")
                r = session.get(f"{base_url}{data['_expandable']['children']}", verify=False)
                print(r.status_code)
                data2 = r.json()
                print(json.dumps(data2, indent=4) )
                for property in  ['page', 'attachment']:
                    print (f"retrieving URL for page property '{property}' {data2['_expandable'][property]}")
                    r = session.get(f"{base_url}{data2['_expandable'][property]}", verify=False)
                    print(r.status_code)
                    data3 = r.json()
                    if property == 'page':
                        for data4 in data3['results']:
                            print (f"Child Page: {data4['title']}, id: {data4['id']}, url: {data4['_links']['self']}")
                            
                            if retrieve_page_details_2(session,base_url,user_base_url,data4['id'],table_list,treewalk_dict) < 0:
                                return -1 
                            
                    if property == 'attachment':

                        for data4 in data3['results']:
                            attach_dict = {}
                            if not 'attachments' in treewalk_dict: 
                                treewalk_dict['attachments'] = []
                            attach_dict['host_document'] = data['title']
                            attach_dict['doc_url'] = f"{user_base_url}/{page_id}"
                            for element in ['title', 'id', 'type','status' ]:
                                attach_dict[element] = data4[element]
                            attach_dict['download_link'] = f"{base_url}{data4['_links']['download']}"
                                    
                            treewalk_dict['attachments'].append(attach_dict)
                            print(f"found attachment:", json.dumps(data4, indent = 4) )
                            increase_attachment_count(treewalk_dict)
        else: 
            print(f"Document id:{page_id}, Title: {data['title']} Does not meet format template (first table does not meet requirements), ignored")
            if not 'unparsed_documents' in treewalk_dict: 
                treewalk_dict['unparsed_documents'] = []
            doc = { "page_id": page_id, "title": data['title'], 'reason': "first table does not meet template requirements", "doc_url":f"{user_base_url}/{page_id}" }
            treewalk_dict['unparsed_documents'].append(doc)  
            increase_unparsed_page_count(treewalk_dict)  
    else: 
        print(f"Document id:{page_id}, Title: {data['title']} Does not meet format template (first table does not exist), ignored")
        if not 'unparsed_documents' in treewalk_dict: 
            treewalk_dict['unparsed_documents'] = []
            
        doc = { "page_id": page_id, "title": data['title'], 'reason': "first table does not exist", "doc_url":f"{user_base_url}/{page_id}" }
        treewalk_dict['unparsed_documents'].append(doc)  
        increase_unparsed_page_count(treewalk_dict) 
    return 0




def main(): 
    parser = argparse.ArgumentParser(description="Confluence-treewalk. Use -c to identify configuration file and -s to override start doc number")
    
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
    print("arguments parsed")
    try: 
        with open(config_file, "rt") as file: 
            config_object = json.load(file)
    except Exception as e:
        print(f"Problem reading {config_file}: {e}")
        exit(1)
    print("config file ingested")
    if 'cookies' in config_object:
        cookies = config_object['cookies']
        print("cookies present")
    else:
        print(f'no cookies in config file: {config_file}')
        exit(1)
    if 'BASE_URL' in config_object: 
        BASE_URL = config_object['BASE_URL']
        print("BASE_URL present")
    else: 
        print("no BASE_URL in config_object!")
        exit(1)
    if 'USER_BASE_URL' in config_object: 
        USER_BASE_URL = config_object['USER_BASE_URL']
        print("USER_BASE_URL present")
    else: 
        print("no USER_BASE_URL in config_object!")
        exit(1)
    if not 'START_DOC' in config_object and not args.startdoc:
        print("no START_DOC in config_object and no startdoc parameter!")
        exit(1)
    else:   
       
        if args.startdoc:
            START_DOC = args.startdoc
            print(f"Taking {START_DOC} as START_DOC from cmd line parameters")
        else:
            START_DOC= config_object['START_DOC']
            print(f"Taking {START_DOC} as START_DOC from config file")
    if not "tables" in config_object:
        print(f"no TABLES entry in config document: {config_file}")
        exit(1)



    # Get local time with system timezone
    now = datetime.now().astimezone()

    # Format full string with timezone name and offset
    start_time_str = now.strftime("%Y-%m-%d %H:%M:%S %Z%z")
    s = requests.Session()
    s.cookies.update(cookies)
    treewalk_dict = { "confluence_requirements_treewalk": {
            "BASE_URL": BASE_URL,
            "USER_BASE_URL": USER_BASE_URL, 
            "INITAL_DOCUMENT": START_DOC,
            "TIME_OF_GENERATION_START": start_time_str }}

    print(f"Starting Treewalk: {treewalk_dict}")
    if retrieve_page_details_2(s, BASE_URL,USER_BASE_URL,START_DOC,config_object['tables'],treewalk_dict) <0: 
        # an error occurred
        print('An error occurred during treewalk - please check logs and ensure you have proper access (check cookies parameter in config)')
    else: 
        # Get local time with system timezone
        now = datetime.now().astimezone()

        # Format full string with timezone name and offset
        end_time_str = now.strftime("%Y-%m-%d %H:%M:%S %Z%z")
        treewalk_dict['confluence_requirements_treewalk']['TIME_OF_GENERATION_END'] = end_time_str
        with open(f"confluence_treewalk_{START_DOC}.json", "w", encoding="utf-8") as file: 
            json.dump(treewalk_dict,file,ensure_ascii=False, indent =4)
        generate_spreadsheet.generate_spreadsheet(f"confluence_treewalk_{START_DOC}.json",config_object)

if __name__ == main():
    main()
