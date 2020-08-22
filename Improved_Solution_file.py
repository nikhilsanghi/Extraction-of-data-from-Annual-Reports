
file_path="/Users/nikhilsanghi/Desktop/Data_Science/NLP/p4/TXT/"
# company_list = ["dufry_ag_2018"
#             ,"dufry_ag_2019"
#             ,"mikron_holding_ag_2018"
#             ,"mikron_holding_ag_2019"
#             ,"allianz_se_2018"
#             ,"allianz_se_2019"
#             ,"eurokai_gmbh_2018"
#             ,"eurokai_gmbh_2019"
#             ,"evolva_holding_ag_2018"
#             ,"evolva_holding_ag_2018"
#             ,"first_derivatives_plc_2018"
#             ,"first_derivatives_plc_2019"
#             ,"aberdeen_asset_management_plc_2018"
#             ,"aberdeen_asset_management_plc_2019"]
company_list = ["icici2020"]


# def pdf_converter(filename):
    # ! pdftotext /Users/nikhilsanghi/Desktop/Data_Science/NLP/p4/PDF/{filename}.pdf /Users/nikhilsanghi/Desktop/Data_Science/NLP/p4/TXT/{filename}.txt   

def fetch_files():
    file_list=[file_path+"{}.txt".format(company_list[i]) for i in range(len(company_list))]
    return file_list

import time
import spacy
import en_core_web_md


start_time = time.time()
nlp = en_core_web_md.load()
nlp.max_length = 1500000
nlp.add_pipe(nlp.create_pipe('sentencizer'))

bad_words=["the Board of Directors","The Board of Directors","THE BOARD OF DIRECTORS"
           ,"Consolidated Financial Statements","Mikron Annual Report","Antonio Gea"
           ,"the Supervisory Board",'',"Group","a000","Company"]

irregularities={ "â\x80\x93toâ\x80\x93mid":"","â\x80\x99s":"", "â\x80\x94":"", "â\x80\x93":""
    ,"â\x80\x99":"","Â£â\x80":"", "â\x80\x9c":"", "â\x80\x9d":"", "Â\xad":"",".\x0c":""
    ,"\x07":"", "ï¬\x81":"fi", "ï¬\x82":"a","\x99000":"","â\x84¢":""
    ,"Ã¡":"a", "Ã\x81":"A", "Ã\xad":"i", "Ã\x89":"EA"
    ,"Ã\xada":"ia", "Â£":"a", "*+":"", "Ã£":"a", "Ã¤":"e"
    ,"¤":"","á":"a","Ã":"A", "Ã¼":"u"
    ,"\n":""}

def get_doc(filename):
    print("Calculating doc...")
    with open(filename,'r',encoding='Latin-1')as f:
        raw_text=f.read()
    doc = nlp(raw_text)
    return doc

def get_company_name(doc):
    print("Calculating company name...")
    org_list=[ent.text for sent in doc.sents for ent in sent.ents if ((ent.label_ == 'ORG') and len(ent.text) > 5)]
    for i, j in irregularities.items():
        org_list_clean = [items.replace(i,j) for items in org_list]
    new_list=[]
    new_list=[item for item in org_list_clean if item not in bad_words]
    comp_name =max(set(new_list), key=new_list.count)
    return comp_name




def get_most_common_name(doc):
    print("Calculating most common name...")
    persons_list=[ent.text for sent in doc.sents for ent in sent.ents if ((ent.label_ == 'PERSON') and len(ent.text) > 5)]

    for i, j in irregularities.items():
        persons_list_clean = [items.replace(i,j) for items in persons_list]
    
    persons_dict={}
    for item in persons_list_clean:
        if item not in bad_words:
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

    for i, j in irregularities.items():
        ceo_list_clean = [items.replace(i,j) for items in ceo_list]

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



file_list=fetch_files()
print(file_list)




for i in range(len(file_list)):
    print("------------------------------------------------------------------------------")
    print("Showing for {}".format(company_list[i]))
    print("------------------------------------------------------------------------------")
    start_time_single_process=time.time()
    data_extracted_dict=get_parameters(get_doc(file_list[i]))
    end_time_single_process=time.time()
    print("------------------------------------------------------------------------------")
    print("time taken : {} secs".format(end_time_single_process-start_time_single_process))
    print(data_extracted_dict)
    print("------------------------------------------------------------------------------")
end_time=time.time()

print("Total time taken : {} secs".format(end_time-start_time))


