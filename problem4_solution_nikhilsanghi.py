import os 

####### Use this to install the required libraries#########
# os.system("pip install -U spacy")
# os.system("python -m spacy download en_core_web_md")
# os.system("conda install -c conda-forge poppler")


######## The code automatically converts pdfs to text file and processes the file for 
######## all the 3 variables and stores them in excel file.
######## create a folder and add all the pdfs of annual reports
######## 1. Add the folder path to the variable pdf_file_path and respectively others.
######## 2. Add the same path to all the 3 variables below if you dont require different folders.

pdf_file_path="/Users/nikhilsanghi/Downloads/Current_Work/nlp/p4/PDF/"
txt_file_path="/Users/nikhilsanghi/Downloads/Current_Work/nlp/p4/TXT/"
output_file_path="/Users/nikhilsanghi/Downloads/Current_Work/nlp/p4/Output/"


import time
import pandas as pd
import en_core_web_md


start_time = time.time()
nlp = en_core_web_md.load()
nlp.max_length = 1500000
nlp.add_pipe(nlp.create_pipe('sentencizer'))

bad_words=["the Board of Directors","The Board of Directors"
           ,"THE BOARD OF DIRECTORS","Consolidated Financial Statements"
           ,"Mikron Annual Report","Antonio Gea","the Supervisory Board"
           ,'',"Group","a000","Company","Committee","theCompany"
           ,"the Audit Committee","Cash Credit","the Annual Report"
           ,"Officer/","TheCompany","Board of Directors"]

def get_company_list(pdf_file_path):
    pdf_file_list=os.listdir(pdf_file_path)
    company_list = [s.replace(".pdf","") for s in pdf_file_list if ".pdf" in s]
    return company_list

def pdf_to_text_converter():
    for i in range(len(company_list)):
        os.system("pdftotext {}{}.pdf {}{}.txt"
                  .format(pdf_file_path,company_list[i]
                          ,txt_file_path,company_list[i]))

def fetch_txt_files():
    txt_file_list=[txt_file_path+"{}.txt".format(company_list[i]) for i in range(len(company_list))]
    return txt_file_list

def get_doc(filename):
    print("Calculating doc...")
    with open(filename,'r',encoding='Latin-1')as f:
        raw_text=f.read()
    doc = nlp(raw_text)
    return doc

def get_company_name(doc):
    print("Calculating company name...")
    org_list=[ent.text for sent in doc.sents for ent in sent.ents if ((ent.label_ == 'ORG') and len(ent.text) > 5)]
    org_list_clean=[item.replace("\n", "")                    
                    .replace("â\x80\x99s","")
                    .replace("â\x80\x93toâ\x80\x93mid","")
                    .replace("â\x80\x94","")
                    .replace("â\x80\x93","")
                    .replace("â\x80\x99","")
                    .replace("â\x80\x9c","")
                    .replace("â\x80\x9d","")
                    .replace("Â\xad","")
                    .replace(".\x0c","")
                    .replace("\x07","")
                    .replace("ï¬\x81","fi")
                    .replace("Ã¼","u")
                    .replace("ï¬\x82","a")
                    .replace("Ã¤","e")
                    .replace("\t\t\t\t","")
                    .replace("Nigrie STPP","")
                    .replace("\x0c"," ") for item in org_list]
   
    org_dict={}
    for item in org_list_clean:
        if item not in org_dict:
            org_dict[item]=1
        else:
            org_dict[item]+=1  
    org_dict_clean=dict((key,value) for key, value in org_dict.items() if key not in bad_words)
    comp_name = max(org_dict_clean, key=org_dict_clean.get)
    return comp_name

def get_most_common_name(doc):
    print("Calculating most common name...")
    persons_list=[ent.text for sent in doc.sents for ent in sent.ents if ((ent.label_ == 'PERSON') and len(ent.text) > 5)]

    persons_list_clean=[item.replace("\n", "")
                        .replace("Nigrie STPP","")
                        .replace(" â\x80\x93","")
                        .replace("â\x80\x99","")
                        .replace("â\x80\x9d","")
                        .replace("â\x80\x9c","")
                        .replace("Ã\xad","i")
                        .replace("Ã\x89","A")
                        .replace("Ã¼","u")
                        .replace("â\x84¢","")
                        .replace("Ã\xada","ia")
                        .replace("Â\xad","")
                        .replace("Ã\x81","A")
                        .replace("\x0c","")
                        .replace("Ã¡","a")
                        .replace("Ã¤","")
                        .replace("*+","")
                        .replace("Â£","a")
                        .replace("Ã£","a") for item in persons_list]
    
    persons_dict={}
    for item in persons_list_clean:
        if item not in persons_dict:
            persons_dict[item]=1
        else:
            persons_dict[item]+=1
    persons_dict_clean=dict((key,value) for key, value in persons_dict.items() if key not in bad_words)
    most_common_name = max(persons_dict_clean, key=persons_dict_clean.get)
    return most_common_name

def get_ceo_name(doc):
    print("Calculating CEO...")
    ceo_list=[]
    for sent in doc.sents:
        for tag in ['ceo', 'chief executive officer']:
            if tag in sent.text.lower():
                for ent in sent.ents:
                    if (ent.label_ == 'PERSON') and len(ent.text) > 5:
                        ceo_list.append(ent.text)
                                         
    if not ceo_list:
       ceo_list=[ent.text for sent in doc.sents for ent in sent.ents if ((ent.label_ == 'PERSON') and len(ent.text) > 5)]

    ceo_list_clean=[item.replace("\n", "")
               .replace("â\x80\x93","")
               .replace("â\x80\x99s","")
               .replace("Ã\xada","ia")
               .replace("\x99000","")
               .replace("Â£â\x80","")
               .replace("*+","")
               .replace("Â£â\x80","")
               .replace("\x81","")
               .replace("Â£","")
               .replace("Ã¤","a")
               .replace("Ã¡","a")
               .replace("á","a")
               .replace("Ã","A")                
               .replace("¤","") for item in ceo_list]

    ceo_dict={}
    for item in ceo_list_clean:
        if item not in ceo_dict:
            ceo_dict[item]=1
        else:
            ceo_dict[item]+=1
    
    ceo_dict_clean=dict((key,value) for key, value in ceo_dict.items() if key not in bad_words)
    ceo_name = max(ceo_dict_clean, key=ceo_dict_clean.get)
    return ceo_name

def get_parameters(doc):
    data_extracted = {}
    data_extracted["Org"] = get_company_name(doc)
    print("Org : {}".format(data_extracted["Org"]))
    data_extracted["Person"] = get_most_common_name(doc)
    print("Person : {}".format(data_extracted["Person"]))
    data_extracted["CEO"] = get_ceo_name(doc)  
    print("Ceo : {}".format(data_extracted["CEO"]))
    return data_extracted

company_list=get_company_list(pdf_file_path)
pdf_to_text_converter()
txt_file_list=fetch_txt_files()

column_names = ["Org", "Person", "CEO","Time Taken(secs)"]
df = pd.DataFrame(columns = column_names)

for i in range(len(txt_file_list)):
    print("----------------------------------------------")
    print("Showing for {}".format(company_list[i]))
    print("----------------------------------------------")
    start_time_single_process=time.time()
    data_extracted_dict=get_parameters(get_doc(txt_file_list[i]))
    end_time_single_process=time.time()
    print("----------------------------------------------")
    print("time taken : {} secs".format(end_time_single_process-start_time_single_process))
    print(data_extracted_dict)
    print("----------------------------------------------")
    print("#   #   #   #   #   #   #   #   #   #   #   # ")
    df2 = pd.DataFrame({ "Org" : '1'
                        ,"Person" : '1'
                        ,"CEO":'1'
                        ,"Time Taken(secs)":[end_time_single_process-start_time_single_process]}
                        ,index=["{}".format(company_list[i])]) 
    for k,v in data_extracted_dict.items():
        df2[k] = v
    df = pd.concat([df, df2])
end_time=time.time()
print("----------------------------------------------")
print("Total time taken : {} secs".format(end_time-start_time))
print("----------------------------------------------")
df.to_excel("{}output.xlsx".format(output_file_path))

