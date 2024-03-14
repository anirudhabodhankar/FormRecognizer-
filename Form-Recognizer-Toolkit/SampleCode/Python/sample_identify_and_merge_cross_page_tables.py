#!/usr/bin/env python
# coding: utf-8

# # Identify cross-page tables based on rules
# 
# This sample demonstrates how to use the output of Layout model and some business rules to identify cross-page tables. Once idenfied, it can be further processed to merge these tables and keep the semantics of a table.
# 
# Depending on your document format, there can be different rules applied to idenfity a cross-page table. This sample shows how to use the following rules to identify cross-page tables:
# 
# - If the 2 tables appear in consecutive pages
# - And there's only page header, page footer or page number beteen them
# - And the tables have the same number of columns
# 
# You can customize the rules based on your scenario.

# ## Prerequisites
# - An Azure AI Document Intelligence resource - follow [this document](https://learn.microsoft.com/azure/ai-services/document-intelligence/create-document-intelligence-resource?view=doc-intel-4.0.0) to create one if you don't have.
# - Get familiar with the output structure of Layout model - complete [this quickstart](https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/quickstarts/get-started-sdks-rest-api?view=doc-intel-4.0.0&pivots=programming-language-python#layout-model) to learn more.

# ## Setup


"""
This code loads environment variables using the `dotenv` library and sets the necessary environment variables for Azure services.
The environment variables are loaded from the `.env` file in the same directory as this notebook.
"""
import os
from dotenv import load_dotenv
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest, ContentFormat, AnalyzeResult
from bs4 import BeautifulSoup

load_dotenv()

endpoint = os.getenv("AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT")
key = os.getenv("AZURE_DOCUMENT_INTELLIGENCE_KEY")


def table_to_html(table):
    cells = sorted(table['cells'], key=lambda x: (x['rowIndex'], x['columnIndex']))

    # Initialize the HTML table
    html = '<table>\n'
    
    # Initialize rowIndex and columnIndex
    current_row = cells[0]['rowIndex']
    for cell in cells:
        # If rowIndex has changed, close the previous row and start a new one
        if cell['rowIndex'] != current_row:
            html += "</tr>"
            
        
        # If rowIndex has changed or it's the first cell, start a new row
        if cell['columnIndex'] == 0 or cell['rowIndex'] != current_row:
            html += "<tr>"
            current_row = cell['rowIndex']
        
        # Add the cell to the row
        tag = "th" if cell.get('kind') == 'columnHeader' else "td"
        content = cell['content'].replace('\\\\', '')  # remove escape sequence
        html += f'<{tag} rowspan="{cell.get("rowSpan", 1)}" colspan="{cell.get("columnSpan", 1)}">{content}</{tag}>'
    
    # Close the last row and the table
    html += "</tr></table>"
    
    return html


# In[120]:


def get_table_page_numbers(table):
    """
    Returns a list of page numbers where the table appears.

    Args:
        table: The table object.

    Returns:
        A list of page numbers where the table appears.
    """
    return [region.page_number for region in table.bounding_regions]


# In[121]:


def get_table_span_offsets(table):
    """
    Calculates the minimum and maximum offsets of a table's spans.

    Args:
        table (Table): The table object containing spans.

    Returns:
        tuple: A tuple containing the minimum and maximum offsets of the table's spans.
    """
    min_offset = table.spans[0].offset
    max_offset = table.spans[0].offset + table.spans[0].length

    for span in table.spans:
        if span.offset < min_offset:
            min_offset = span.offset
        if span.offset + span.length > max_offset:
            max_offset = span.offset + span.length

    return min_offset, max_offset


# In[131]:


def prepare_html_tables(tables):
    """
    Converts all tables to html format. 

    Parameters:
    tables (list): A list of tables.

    Returns:
    list: A list of tables with meta data 
    """
    html_tables = []

    for table_idx, table in enumerate(tables):
        min_offset, max_offset = get_table_span_offsets(table)
        table_page = get_table_page_numbers(table)
        data_html = table_to_html(table)
        
        soup = BeautifulSoup(data_html, 'html.parser')
        souptable = soup.find('table')
        header_rows = souptable.find_all(lambda tag: tag.name == 'tr' and tag.find('th'))
        header_rows_count = len(header_rows) if header_rows else 0

        header_cells = soup.find_all('th')
        header_cell_texts = []
        for cell in header_cells:
            header_cell_texts.append(cell.get_text())

        header_cell_texts.sort()
        
        html_table = {
            "min_offset": min_offset, 
            "max_offset": max_offset, 
            "table_page": table_page,
            "column_count" : table.column_count,
            "header_row_count" : header_rows_count,
            "header_text" : header_cell_texts,
            "non_header_row_count" : table.row_count - header_rows_count,
            "content": data_html
        }
        
        html_tables.append(html_table)
        
    return html_tables


# In[132]:


def find_merge_table_candidates(html_lables):
    """
    Finds the merge table candidates based on the given list of tables.

    Parameters:
    tables (list): A list of tables.

    Returns:
    list: A list of merge table candidates, where each candidate is a dictionary with keys:
          - pre_table_idx: The index of the first candidate table to be merged (the other table to be merged is the next one).
          - start: The start offset of the 2nd candidate table.
          - end: The end offset of the 1st candidate table.
    """
    merge_tables_candidates = []
    pre_table_idx = -1
    pre_table_page = -1
    pre_max_offset = 0

    for table_idx, table in enumerate(html_lables):
        min_offset= table["min_offset"] 
        max_offset = table["max_offset"]
        table_page = min(table["table_page"])
        
        # If there is a table on the next page, it is a candidate for merging with the previous table.
        if table_page == pre_table_page + 1:
            pre_table = {
                "pre_table_idx": pre_table_idx, 
                "start": pre_max_offset, 
                "end": min_offset
            }

            merge_tables_candidates.append(pre_table)
        
        #print(f"Table {table_idx} has offset range: {min_offset} - {max_offset} on page {table_page}")

        pre_table_idx = table_idx
        pre_table_page = max(table["table_page"])
        pre_max_offset = max_offset

    return merge_tables_candidates


# In[133]:


def check_paragraph_presence(paragraphs, start, end):
    """
    Checks if there is a paragraph within the specified range that is not a page header, page footer, or page number. If this were the case, the table would not be a merge table candidate.

    Args:
        paragraphs (list): List of paragraphs to check.
        start (int): Start offset of the range.
        end (int): End offset of the range.

    Returns:
        bool: True if a paragraph is found within the range that meets the conditions, False otherwise.
    """
    for paragraph in paragraphs:
        for span in paragraph.spans:
            if span.offset > start and span.offset < end:
                # The logic role of a parapgaph is used to idenfiy if it's page header, page footer, page number, title, section heading, etc. Learn more: https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/concept-layout?view=doc-intel-4.0.0#document-layout-analysis
                if not hasattr(paragraph, 'role'):
                    return True
                elif hasattr(paragraph, 'role') and paragraph.role not in ["pageHeader", "pageFooter", "pageNumber"]:
                    return True
    return False


# In[134]:


def merge_tables_colum_wise(table1_html, table2_html):
    # Parse the tables with BeautifulSoup
    soup1 = BeautifulSoup(table1_html['content'], 'html.parser')
    soup2 = BeautifulSoup(table2_html['content'], 'html.parser')
    
    # Find the tables
    table1 = soup1.find('table')
    table2 = soup2.find('table')
    
    # For each row in table1, find the corresponding row in table2 (by index) and append its columns
    for row1, row2 in zip(table1.find_all('tr'), table2.find_all('tr')):
        for column in row2.find_all(['td', 'th']):
            row1.append(column)
    
    # Print the modified table1 as a string
    merged_table_html = {
            "min_offset": min(table1_html["min_offset"], table2_html["min_offset"]),
            "max_offset": max(table1_html["max_offset"], table2_html["max_offset"]),
            "table_page": table1_html["table_page"] + table2_html["table_page"],
            "column_count" : table1_html["column_count"] + table2_html["column_count"],
            "header_row_count" : table1_html["header_row_count"],
            "header_text" : table1_html["header_text"] + table2_html["header_text"],
            "non_header_row_count" : table1_html["non_header_row_count"],
            "content": str(table1)
        }
    merged_table_html["table_page"].sort()
    merged_table_html["header_text"].sort()
    
    #print(merged_table_html)
    return merged_table_html



def merge_tables_row_wise(table1_html, table2_html):
    # Parse the tables with BeautifulSoup
    soup1 = BeautifulSoup(table1_html['content'], 'html.parser')
    soup2 = BeautifulSoup(table2_html['content'], 'html.parser')
    
    # Find the tables
    table1 = soup1.find('table')
    table2 = soup2.find('table')
    
    # Check if second table has headers (th elements)
    header_rows = table2.find_all(lambda tag: tag.name == 'tr' and tag.find('th'))
    
    # Skip headers if present
    start_index = len(header_rows) if header_rows else 0
        
    # Get all rows in second table, excluding the header rows
    rows = table2.find_all('tr')[start_index:]
    
    # Append each row from the second table to the first
    for row in rows:
        table1.append(row)
    
    # The merged table
    merged_table_html = {
            "min_offset": min(table1_html["min_offset"], table2_html["min_offset"]),
            "max_offset": max(table1_html["max_offset"], table2_html["max_offset"]),
            "table_page": table1_html["table_page"] + table2_html["table_page"],
            "column_count" : table1_html["column_count"],
            "header_row_count" : table1_html["header_row_count"],
            "header_text" : table1_html["header_text"],
            "non_header_row_count" : table1_html["non_header_row_count"] + table2_html["non_header_row_count"],
            "content": str(table1)
        }
    merged_table_html["table_page"].sort()
    
    #print(merged_table_html)
    return merged_table_html


# In[135]:


def check_and_merge_column_wise(paragraphs, html_tables, merge_tables_candidates):
    for i, candidate in enumerate(merge_tables_candidates):
        table_idx = candidate["pre_table_idx"]
        start = candidate["start"]
        end = candidate["end"]
        has_paragraph = check_paragraph_presence(paragraphs, start, end)
        has_paragraph= False
                
        table1 = html_tables[table_idx];
        table2 = html_tables[table_idx + 1];
        
        # If there is no paragraph within the range and the columns of the tables match, merge the tables.
        if not has_paragraph and table1["header_row_count"] == table2["header_row_count"] and table1["non_header_row_count"] == table2["non_header_row_count"]:
            #print(f"Merge table: {table_idx} and {table_idx + 1}")
            html_tables[table_idx + 1] = merge_tables_colum_wise(table1, table2)
            html_tables[table_idx] = None
            #print("----------------------------------------")
        


# In[136]:


def check_and_merge_row_wise(paragraphs, html_tables, merge_tables_candidates):
    for i, candidate in enumerate(merge_tables_candidates):
        table_idx = candidate["pre_table_idx"]
        start = candidate["start"]
        end = candidate["end"]
        has_paragraph = check_paragraph_presence(paragraphs, start, end)
        has_paragraph= False
                
        table1 = html_tables[table_idx];
        table2 = html_tables[table_idx + 1];
        
        # If there is no paragraph within the range and the columns of the tables match, merge the tables.
        if (not has_paragraph 
            and table1["column_count"] == table2["column_count"] 
            and ((table1["header_row_count"] <=0 and table2["header_row_count"] <= 0) 
                or (table1["header_row_count"] == table2["header_row_count"] and table1["header_text"] == table2["header_text"] ))
           ):
            #print(f"Merge table: {table_idx} and {table_idx + 1}")
            html_tables[table_idx + 1] = merge_tables_row_wise(table1, table2)
            html_tables[table_idx] = None
            #print("----------------------------------------")


# In[139]:

def chunks(list, start_index, chunk_size):
    """Yield successive n-sized chunks from lst."""
    for i in range(start_index, len(list), chunk_size):
        yield list[i:i + chunk_size]
        
def split_table_with_headers(table_html, table_max_rows):
    soup = BeautifulSoup(table_html['content'], 'html.parser')  
    # Find the tables
    table = soup.find('table')

    # Check if second table has headers (th elements)
    header_rows = table.find_all(lambda tag: tag.name == 'tr' and tag.find('th'))
    non_header_rows = table.find_all(lambda tag: tag.name == 'tr' and not tag.find('th'))

    header_string = str(header_rows)[1:-1].replace('</tr>, <tr>', '</tr><tr>')  if header_rows else ""
    new_html_table_list = []

    for chunk in chunks(non_header_rows, 0, table_max_rows):
        new_table = "<table>"
        new_table = new_table + header_string
        new_table = new_table + str(chunk)[1:-1].replace('</tr>, <tr>', '</tr><tr>') if header_rows else ""
        new_table = new_table + '</table>'
        

        new_html_table = {
            "min_offset": table_html["min_offset"],
            "max_offset": table_html["max_offset"],
            "table_page": table_html["table_page"],
            "column_count" : table_html["column_count"],
            "header_row_count" : table_html["header_row_count"],
            "header_text" : table_html["header_text"],
            "non_header_row_count" : table_html["non_header_row_count"],
            "content": new_table
        }
        
        new_html_table_list.append(new_html_table)
        #print(new_html_table)

    return new_html_table_list


def identify_cross_page_tables(result, table_max_rows = 5):
    """
    Identifies and merges tables that span across multiple pages in a document.
    
    Returns:
    None
    """

    html_tables = prepare_html_tables(result.tables)

    #First try merge all the tables that are split columnwise
    merge_tables_candidates = find_merge_table_candidates(html_tables)
    #print(merge_tables_candidates)
    #print("----------------------------------------")
    check_and_merge_column_wise(result.paragraphs, html_tables, merge_tables_candidates)
    html_tables = [x for x in html_tables if x is not None]

    #Now try merge all the tables that are split rowwise
    merge_tables_candidates = find_merge_table_candidates(html_tables)
    #print(merge_tables_candidates)
    #print("----------------------------------------")
    check_and_merge_row_wise(result.paragraphs, html_tables, merge_tables_candidates)
    html_tables = [x for x in html_tables if x is not None]

    final_html_tables =[]
    for table in html_tables:  
        final_html_tables.extend(split_table_with_headers(table, table_max_rows))

    #print(html_tables)
    return(final_html_tables)



