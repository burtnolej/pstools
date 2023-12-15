from pptx import Presentation
from openpyxl import Workbook
from  openpyxl import open as openxl
import win32com.client
from os import mkdir, path
from shutil import copy
from types import NoneType

#prs = Presentation()
#prs.save('Master_Whitepaper.pptx')


def get_range(wb,sheetname,address,rangename):
    ws = wb[sheetname]
    globals()["_".join([sheetname.lower(),rangename,"range"])] = ws[address]
             
def get_map(xlmap,row,col,data):    
    
    for j in range(0,len(col[0])):
        _fields=[]
        file_type=col[0][j].value
        for i in range(0,len(row)):
            if data[i][j].value == "Y":
                _fields.append(row[i][0].value)
        xlmap[file_type]=_fields
        
def get_store(store):    
    for i in range(0,len(data_rowtitle_range)):
        file_type=data_filetype_range[i][0].value
        for j in range(0,len(data_colheader_range[0])):
            article_name=data_colheader_range[0][j].value
            field_name=data_rowtitle_range[i][0].value
            field_value=data_data_range[i][j].value
            
            if article_name not in store:
                store[article_name]={}
                
            store[article_name][field_name]=field_value


def _create_article_dir(topdir,_article):
    currentdir=path.join(topdir,_article)
    if path.isdir(currentdir)==False:
        mkdir(currentdir)
    return currentdir
    
def _get_default_name(filename,defaultname="thumbnail"):
    ext = path.splitext(filename)
    bname = path.basename(filename)
    _new_file_name=defaultname+ ext[1]
    return _new_file_name
    
def generate_config_files(doc_types,store,file_map):
    for doc_type in doc_types:
        print(doc_type)
        
        for _article in store.keys():
            currentdir=_create_article_dir(topdir,_article)
            
            outputfile=path.join(currentdir,doc_type+".txt")
            with open(outputfile,"w+") as f:
                for _field in file_map[doc_type]:
                    try:
                        if _field not in store[_article]:
                            print("mandatory field ",_field," not found for article ",_article," in file type ",doc_type)
                        else:
                            f.write(_field+"*"+store[_article][_field]+"\n")
                    except Exception as err:
                        print(type(err),_article,doc_type,_field)
                    

def process_artefacts(store,topdir):
    for _article in store.keys():
        _image_file=store[_article]["_image"]
        _content_file=store[_article]["_content"]
        if isinstance(_image_file,NoneType) == True:
            print("_image_file"," cannot be None for article ",_article)
            exit()
            
        if path.isfile(_image_file)==True:

            _new_file_name=_get_default_name(_image_file)
            currentdir=_create_article_dir(topdir,_article)
            store[_article]["_image"]=_new_file_name
            #print ("copy " + _image_file + " to " + _new_file_name + " for " + _article)
            copy(_image_file,path.join(topdir,_article,_new_file_name))
            
        if path.isfile(_content_file)==True:
            _new_file_name=_get_default_name(_content_file,"article")
            currentdir=_create_article_dir(topdir,_article)
            store[_article]["_content"]=_new_file_name
            #print ("copy " + _content_file + " to " + _new_file_name + " for " + _article)
            copy(_content_file,path.join(topdir,_article,_new_file_name))
        
def generate_inclusion_files(inclusion_map, target_folder,contenttype_map,visibility_map,latest_map):    

    for _outputtype in inclusion_map.keys():
        with open(path.join(target_folder,_outputtype+"_inclusion.txt"),"w+") as f:
            _latest="False"
            for _element in inclusion_map[_outputtype]:
                if latest_map[_element]=="Y":
                    _latest="True"
                    
                f.write(_element+"," + contenttype_map[_element]+"," + str(visibility_map[_element])+"," + _latest + "\n")
                _latest="False"
    inclusion_map["latest"]=[]
     
    for key in latest_map.keys():
        if latest_map[key]=="Y":
            inclusion_map["latest"].append(key)
           
def get_dict_from_ranges(_dict,xlrange_key,xlrange_value,row=True):   
    if row==True:
        for i in range(len(xlrange_key)):
            _dict[xlrange_key[i][0].value]=xlrange_value[i][0].value
    return _dict

doc_types=["snippet_docs","snippet","article","teaser"]
file_map={}
store={}
contenttype_map={}
visibility_map={}
inclusion_map={}
latest_map={}
topdir="articles"

for _doc_type in doc_types:
    file_map[_doc_type]=[]

wb = openxl("website_map.xlsx")
ws_data = wb["DATA"]
ws_map = wb["MAP"]
    
get_range(wb,"DATA","C1:AM1","colheader")
get_range(wb,"DATA","A2:A13","rowtitle")
get_range(wb,"DATA","B2:B13","filetype")
get_range(wb,"DATA","C2:AM13","data")

get_range(wb,"MAP","B1:E1","colheader")
get_range(wb,"MAP","A2:A13","rowtitle")
get_range(wb,"MAP","B2:E13","data")

get_range(wb,"INCLUSION","B1:D1","colheader")
get_range(wb,"INCLUSION","A2:A38","rowtitle")
get_range(wb,"INCLUSION","B2:D38","data")

get_range(wb,"CONTENTTYPE","A2:A38","key")
get_range(wb,"CONTENTTYPE","B2:B38","value")

get_range(wb,"VISIBILITY","A2:A38","key")
get_range(wb,"VISIBILITY","B2:B38","value")

get_range(wb,"LATEST","A2:A38","key")
get_range(wb,"LATEST","B2:B38","value")

get_dict_from_ranges(contenttype_map,contenttype_key_range,contenttype_value_range)
get_dict_from_ranges(visibility_map,visibility_key_range,visibility_value_range)
get_dict_from_ranges(latest_map,latest_key_range,latest_value_range)

get_store(store)
get_map(file_map,map_rowtitle_range,map_colheader_range,map_data_range)
get_map(inclusion_map,inclusion_rowtitle_range,inclusion_colheader_range,inclusion_data_range)


generate_inclusion_files(inclusion_map,"inclusion",contenttype_map,visibility_map,latest_map)
process_artefacts(store,topdir)
generate_config_files(doc_types,store,file_map)


#PptApp = win32com.client.Dispatch("Powerpoint.Application")
#PptApp.Visible = True
#z= excelrange.Copy()
#PPtPresentation = PptApp.Presentations.Open(r'C:\Users\burtn\Development\py\Master_Whitepaper.pptx')
#pptSlide = PPtPresentation.Slides.Add(1,11)
#pptSlide.Title.Characters.Text ='Metrics'

#title = pptSlide.Shapes.Title
#title.Text ='Metrics Summary'
#pptSlide.Shapes.PasteSpecial(z)
#PPtPresentation.SaveAs()