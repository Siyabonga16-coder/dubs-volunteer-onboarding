from docx import Document
import csv


doc = Document("path of the actual document ")

for para in doc.paragraphs:#now printing all the texts contained in the file.
    print(para.text)
#this step act as  a discovery step to see format of the document
volunteer_data = {
    "Name": "",
    "Surname": "",
    "DateOfBirth": "",
    "Cellphone": "",
    "Email": "",
    
    # waiting for the form format to come with the form created
    
}
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
        

    # wiating for another atributes to come with designers
    
