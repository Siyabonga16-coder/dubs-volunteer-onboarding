from docx import Document
import csv
import os

input_folder = "templates/"
output_folder = "output/volunteer.csv"

headers = ["Name", "Surname", "DateOfBirth", "Cellphone", "Email"]# add them if all atrtributes are there in the form
file_exists = os.path.isfile(output_folder) # this is to check if the file already exists, if it does we will not write the header again, if it doesn't we will write the header.
with open(output_folder, mode='a', newline='',encoding='utf-8') as file:#opening the csv file in write mode, if it doesn't exist it will be created, if it exists it will be overwritten
    writer = csv.DictWriter(file,fieldnames=headers)
    if not file_exists:
        writer.writeheader()
    
    
    for filename in os.listdir(input_folder):
        if filename.endswith(".docx"):
            doc_path = os.path.join(input_folder, filename)
            doc = Document(doc_path)
            
            #  I will remove this after the code is tested, this is just to see the format of the document and how the data is structured, we will use this information to extract the data correctly.
            for para in doc.paragraphs:#now printing all the texts contained in the file.
                print(para.text)
            #this step act as  a discovery step to see format
            volunteer_data = {key: "" for key in headers}# this is to create a dict with all the keys and empty values, we will fill it later
            
            

            for para in doc.paragraphs:
                text = para.text.strip()
                if text.startswith("Name:"):
                    volunteer_data["Name"]= text.replace("Name:","").strip()
                elif text.startswith("Surname:"):
                    volunteer_data["Surname"]=text.replace("Surname:","").strip()
                elif text.startswith("Date of Birth:"):
                    volunteer_data["DateOfBirth"] = text.replace("Date of Birth:", "").strip()
                elif text.startswith("Cellphone:"):
                    volunteer_data["Cellphone"] = text.replace("Cellphone:", "").strip()
                elif text.startswith("Email:"):
                    volunteer_data["Email"] = text.replace("Email:", "").strip()
        
               # wiating for another additinal attributes to be added in the form, we can add more elif statements here to extract those attributes as well.
          

