import collections
import os
import pandas as pd
import json

def load_file_path():
    with open("config.json","rb") as f:
        data=json.load(f)
    print(data)
    return (data["result_dir"], data["batch_dir"])        


def extract_resust(result_dir,batch_dir):
    """
    Extracts and organizes data from Excel files in a specified directory.

    This function reads Excel files from the given directory, processes data from
    a specified sheet, and organizes it into a structured dictionary format. The
    data is expected to contain attributes and their values, which are then mapped
    to corresponding pages and linked to HTML file paths.

    Parameters:
    -----------
    result_dir : str
        The directory containing the Excel files to be processed.
    result_file : str
        The file name pattern for result files within `result_dir`. This parameter 
        is reassigned within the function to represent each file being processed.
    batch_dir : str
        The directory where batch files are located. Used to construct file paths for results.

    Returns:
    --------
    final_result : dict
        A dictionary where each key is the base file name (derived from the original 
        result file name), and each value is another dictionary with page numbers as 
        keys and lists of tuples as values. Each tuple contains an attribute and its 
        corresponding extracted value in that page number. Additionally, an HTML file 
        path corresponding to each page is included in the list.
    """
    files = os.listdir(result_dir)
    final_result = collections.defaultdict(dict)
    for file in files:
        result_file=file
        base_file=result_file.split("_")[0]
        visualize_data_path = os.path.join(result_dir, result_file)
        data_xls = pd.read_excel(visualize_data_path, "Sheet1", index_col=None)
        # print(data_xls.columns)
        attribute_column = data_xls.columns.to_list()[0]
        value_column = data_xls.columns.to_list()[1]
        #page_column = data_xls.columns.to_list()[2]
        #print(page_column)
        attribute = data_xls[attribute_column].to_list()
        value = data_xls[value_column].to_list()
        #page = data_xls[page_column].to_list()
        # print(page)
        result = collections.defaultdict(list)
        # print(value)
        for ix, page_attribute in enumerate(value):
                attribute_dict=eval(page_attribute)
                current_attribute=attribute[ix]
                for i in attribute_dict.keys():
                    result[i].append((current_attribute,attribute_dict[i]))
        for page_no in result.keys():
            result[page_no].append(os.path.join(batch_dir,base_file,f"{base_file}.pdf-{page_no}.html"))
        final_result[base_file]=result
    return final_result
    
if __name__=="__main__":
    result_dir,batch_dir=load_file_path()
    final_result=extract_resust(result_dir,batch_dir)
    try:
        with open(final_result["page0"][-1],"rb") as f:
            print(f)
    except:
        print("Failed top open file")