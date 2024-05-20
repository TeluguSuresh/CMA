from tkinter import filedialog
from pdf2image import convert_from_path
from PyPDF2 import PdfReader
import os
import pytesseract
import re
import pandas as pd
import datetime
from concurrent.futures import ThreadPoolExecutor

# Set configuration
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
poppler_path = r'C:\Program Files\poppler-23.01.0\Library\bin'

# Extracting PDF Pages to Images
def convert_pdf_to_images(pdf_path, poppler_path):
    with open(pdf_path, "rb") as file:
        pdf = PdfReader(file)
        total_pages = len(pdf.pages)
        print("PDF Total pages Count", total_pages)

    images = []
    if total_pages > 3:
        images = convert_from_path(pdf_path, first_page=1, last_page=total_pages-1, fmt="jpeg", poppler_path=poppler_path)
    
    print("Total Images To Be Processed is ",len(images))
    
    return images

def process_image_and_ocr(cropped_images):

    extracted_text = []

    for i, CropedImage in enumerate(cropped_images):
        # Performing OCR on the cropped image
        # text = pytesseract.image_to_string(CropedImage, lang='eng')
        text = pytesseract.image_to_string(CropedImage, lang='eng', config='--psm 6')
        extracted_text.append(text)

    return extracted_text

def process_pdf_file(file_path, output_directory):
    try:
            
        print("Processing PDF FileName !:", file_path)

        # Record starting time
        start_time = datetime.datetime.now()
        print("Program Execution Starting Time!",start_time)

        # Convert PDF Pages to images
        images = convert_pdf_to_images(file_path, poppler_path)

        extracted_text_list=[]
        cropped_images = []

        # Data extraction and formatting to Dictionary
        Data_Dict = {
            "Voterids": [],
            "Name": [],
            "GaurdianName": [],
            "House Number": [],
            "Age": [],
            "Gender": [],
            "MainTown/Village": [],
            "PoliceStation": [],
            "Mandal": [],
            "RevenueDivision": [],
            "District": [],
            "PinCode": [],
            "AssemblyConstituency": [],
            "ParliamentaryConstituency": [],
            "PollingStation": [],
            }
        
        # croping custom cells from the image
        for img_index, img in enumerate(images):

            #  index==0 is the First Image/Page From the PDF to Get The Required Information From The Image
            if img_index==0:
                # croping custom cells from the first image
                cropped_image = img.crop((70, 115, 1580, 2110))

                # Coordinates for custom cells
                custom_cells = [
                    (0, 0, 1190, 85),
                    (0, 85, 1400, 165),
                    (645, 600, 1500, 1150),
                    (0, 1210, 820, 1450),
                    ]
                # Coordinates for custom cells
                custom_cells = [
                    (0, 0, 1190, 85),
                    (0, 85, 1400, 165),
                    (645, 600, 1500, 1150),
                    (0, 1210, 820, 1450),
                    ]
                
                # Crop the custom cells and save each one
                for i, (left, upper, right, lower) in enumerate(custom_cells):
                    # Crop the current cell
                    cell = cropped_image.crop((left, upper, right, lower))
                    
                    output_path = os.path.join(output_directory, f"CroppedImage{img_index}_{i}.jpg")
                    # cell.save(output_path)
                    cropped_images.append(cell)
            
                # Perform OCR on the cropped images and all croped images data extraction
                extracted_text_list.extend(process_image_and_ocr(cropped_images))

                # print(extracted_text_list)
                for indexes,i in enumerate(extracted_text_list):
                    if indexes==0:
                        print("Index is 0")
                        i=i.replace("\n","").replace("No. Name","\nNo. Name").replace("Status of Assembly Constituency :","\nAssembly Constituency")
                        i=i.replace("Part","").replace("number","")
                        # print(i)
                        # i=re.sub('[^A-Za-z0-9 ]+', '', i)
                        SplitData=i.split("\n")
                        for strings in SplitData:
                            if strings.startswith("Assembly Constituency"):
                                    strings=re.sub('[^A-Za-z0-9 ()-]+', '', strings)
                                    print(strings)
                                    strings=strings.replace("Assembly Constituency","")
                                    Data_Dict["AssemblyConstituency"].append(strings.strip())

                    if indexes==1:
                        print("Index is 1")
                        # print(i)
                        i=i.replace("is located :","\nParliamentary Constituency").replace("Part","").replace("number","")
                        SplitData=i.split("\n")
                        for strings in SplitData:
                         if strings.startswith("Parliamentary Constituency"):
                                    strings=re.sub('[^A-Za-z0-9 ()-]+', '', strings)
                                    print(strings)
                                    strings=strings.replace("Parliamentary Constituency","")
                                    Data_Dict["ParliamentaryConstituency"].append(strings.strip())
                                    
                    # if indexes==2:
                    #     print("Index is 2")
                    #     i=i.replace("\n"," ")
                    #     i=i.replace("Main Town","\nMain Town").replace("Police Station","\nPolice Station").replace("Mandal","\nMandal")
                    #     i=i.replace("Revenue Division","\nRevenue Division").replace("District","\nDistrict").replace("Pin Code","\nPin Code")
                        
                    #     # print(i)
                    #     SplitData=i.split("\n")
                    #     # SplitData=[x.strip() for x in SplitData if x.strip()]
                    #     # print(SplitData)

                    #     for strings in SplitData:
                    #         # print(strings)
                    #         if strings.startswith("Main Town"):
                    #                 strings=re.sub('[^A-Za-z0-9 /]+', '', strings)
                    #                 print(strings)
                    #                 strings=strings.replace("Main Town/Village","").replace(":","").replace(">","")
                    #                 Data_Dict["MainTown/Village"].append(strings.strip())

                    #         if strings.startswith("Police Station"):
                    #                 strings=re.sub('[^A-Za-z0-9 -]+', '', strings)
                    #                 print(strings)
                    #                 strings=strings.replace("Police Station","").replace(":","").replace(">","")
                    #                 Data_Dict["PoliceStation"].append(strings.strip())

                    #         if strings.startswith("Mandal"):
                    #                 strings=re.sub('[^A-Za-z0-9 -]+', '', strings)
                    #                 print(strings)
                    #                 strings=strings.replace("Mandal","").replace(":","").replace(">","")
                    #                 Data_Dict["Mandal"].append(strings.strip())

                    #         if strings.startswith("Revenue Division"):
                    #                 strings=re.sub('[^A-Za-z0-9 -]+', '', strings)
                    #                 print(strings)
                    #                 strings=strings.replace("Revenue Division","").replace(":","").replace(">","")
                    #                 Data_Dict["RevenueDivision"].append(strings.strip())

                    #         if strings.startswith("District"):
                    #                 strings=re.sub('[^A-Za-z0-9 -]+', '', strings)
                    #                 print(strings)
                    #                 strings=strings.replace("District","").replace(":","").replace(">","")
                    #                 Data_Dict["District"].append(strings.strip())

                    #         if strings.startswith("Pin Code"):
                    #                 strings=re.sub('[^A-Za-z0-9 -]+', '', strings)
                    #                 print(strings)
                    #                 strings=strings.replace("Pin Code","")
                    #                 Data_Dict["PinCode"].append(strings.strip())
                    if indexes==2:
                        print("Index is 2")
                        # print(i)
                        # i=i.replace("\n"," ")
                        # i=i.replace("Main Town","\nMain Town").replace("Police Station","\nPolice Station").replace("Mandal","\nMandal")
                        # i=i.replace("Revenue Division","\nRevenue Division").replace("District","\nDistrict").replace("Pin Code","\nPin Code").strip()
                        
                        i=i.replace("\x0c","")
                        i=i.replace("Main Town","\n").replace("Police Station","\n").replace("Mandal","\n").replace("/Village","\n").replace(">","\n")
                        i=i.replace("Revenue Division","\n").replace("District","\n").replace("Pin Code","\n").replace(":","\n").strip()
                        
                        # print(i)
                        SplitData=i.split("\n")
                        
                        # print(SplitData)
                        Not_Empty_Data = [x.strip() for x in SplitData if x.strip()]
                        print(Not_Empty_Data)
                        # Not_Empty_Data[0]="Main Town "+Not_Empty_Data[-6]
                        Not_Empty_Data[0]="Main Town VillageIsInTelugu"
                        Not_Empty_Data[1]="Police Station "+Not_Empty_Data[-5]
                        Not_Empty_Data[2]="Mandal "+Not_Empty_Data[-4]
                        Not_Empty_Data[3]="Revenue Division "+Not_Empty_Data[-3]
                        Not_Empty_Data[4]="District "+Not_Empty_Data[-2]
                        Not_Empty_Data[5]="Pin Code "+Not_Empty_Data[-1]
                        # Not_Empty_Data[0]="Main Town "+Not_Empty_Data[6]

                        for strings in Not_Empty_Data:
                            # print(strings)


                            # if  len(strings)!=0:
                            #     print(len(strings))
                            #     print(strings)
                            if strings.startswith("Main Town"):
                                    strings=re.sub('[^A-Za-z0-9 /]+', '', strings)
                                    print(strings)
                                    strings=strings.replace("Main Town/Village","").replace(":","").replace(">","")
                                    Data_Dict["MainTown/Village"].append(strings.strip())

                            if strings.startswith("Police Station"):
                                    strings=re.sub('[^A-Za-z0-9 -]+', '', strings)
                                    print(strings)
                                    strings=strings.replace("Police Station","").replace(":","").replace(">","")
                                    Data_Dict["PoliceStation"].append(strings.strip())

                            if strings.startswith("Mandal"):
                                    strings=re.sub('[^A-Za-z0-9 -]+', '', strings)
                                    print(strings)
                                    strings=strings.replace("Mandal","").replace(":","").replace(">","")
                                    Data_Dict["Mandal"].append(strings.strip())

                            if strings.startswith("Revenue Division"):
                                    strings=re.sub('[^A-Za-z0-9 -]+', '', strings)
                                    print(strings)
                                    strings=strings.replace("Revenue Division","").replace(":","").replace(">","")
                                    Data_Dict["RevenueDivision"].append(strings.strip())

                            if strings.startswith("District"):
                                    strings=re.sub('[^A-Za-z0-9 -]+', '', strings)
                                    print(strings)
                                    strings=strings.replace("District","").replace(":","").replace(">","")
                                    Data_Dict["District"].append(strings.strip())

                            if strings.startswith("Pin Code"):
                                    strings=re.sub('[^A-Za-z0-9 -]+', '', strings)
                                    print(strings)
                                    strings=strings.replace("Pin Code","")
                                    Data_Dict["PinCode"].append(strings.strip())


                    if indexes==3:
                        print("Index is 3")
                        i=i.replace("\n","").replace("Name of Polling Station","\nName of Polling Station")
                        i=i.replace("Address of Polling Station","\nAddress of Polling Station")

                        SplitData=i.split("\n")    
                        for Polling in SplitData:
                            # print(Polling)
                            if Polling.startswith("Name of Polling Station"):
                                Polling=re.sub('[^A-Za-z0-9 -]+', '', Polling)
                                print(Polling)
                                Polling=Polling.replace("Name of Polling Station","").replace(":","").strip()
                                Data_Dict["PollingStation"].append(Polling)

                
            # img_index>1 is ignoring the second Image/Page From The PDF And Performing The Other Process After The Second Page
            if img_index>1:
                # cropped_image = img.crop((70, 115, 1580, 2110))
                cropped_image = img.crop((70, 115, 1580, 2140))

                # # Coordinates for custom cells
                # custom_cells = [
                #     (0, 0, 500, 210),     (500, 0, 1000, 210),     (1000, 0, 1500, 210),
                #     (0, 190, 500, 410),   (500, 190, 1000, 410),   (1000, 190, 1500, 410),
                #     (0, 390, 500, 610),   (500, 390, 1000, 610),   (1000, 390, 1500, 610),
                #     (0, 590, 500, 810),   (500, 590, 1000, 810),   (1000, 590, 1500, 810),
                #     (0, 790, 500, 1010),  (500, 790, 1000, 1010),  (1000, 790, 1500, 1010),
                #     (0, 990, 500, 1210),  (500, 990, 1000, 1210),  (1000, 990, 1500, 1210),
                #     (0, 1190, 500, 1410), (500, 1190, 1000, 1410), (1000, 1190, 1500, 1410),
                #     (0, 1390, 500, 1610), (500, 1390, 1000, 1610), (1000, 1390, 1500, 1610),
                #     (0, 1590, 500, 1810), (500, 1590, 1000, 1810), (1000, 1590, 1500, 1810),
                #     (0, 1790, 500, 2010), (500, 1790, 1000, 2010), (1000, 1790, 1500, 2010),
                #     ]
                # Coordinates for custom cells
                custom_cells = [
                    (0, 0, 500, 220),     (500, 0, 1000, 220),     (1000, 0, 1500, 220),
                    (0, 190, 500, 420),   (500, 190, 1000, 420),   (1000, 190, 1500, 420),
                    (0, 390, 500, 620),   (500, 390, 1000, 620),   (1000, 390, 1500, 620),
                    (0, 590, 500, 820),   (500, 590, 1000, 820),   (1000, 590, 1500, 820),
                    (0, 790, 500, 1020),  (500, 790, 1000, 1020),  (1000, 790, 1500, 1020),
                    (0, 990, 500, 1220),  (500, 990, 1000, 1220),  (1000, 990, 1500, 1220),
                    (0, 1190, 500, 1420), (500, 1190, 1000, 1420), (1000, 1190, 1500, 1420),
                    (0, 1390, 500, 1620), (500, 1390, 1000, 1620), (1000, 1390, 1500, 1620),
                    (0, 1590, 500, 1820), (500, 1590, 1000, 1820), (1000, 1590, 1500, 1820),
                    (0, 1790, 500, 2020), (500, 1790, 1000, 2020), (1000, 1790, 1500, 2020),
                    ]

                # Crop the custom cells and save each one
                for i, (left, upper, right, lower) in enumerate(custom_cells):
                    # Crop the current cell
                    cell = cropped_image.crop((left, upper, right, lower))
                    
                    output_path = os.path.join(output_directory, f"CroppedImage{img_index}_{i}.jpg")
                    # cell.save(output_path)
                    cropped_images.append(cell)
                print(f"{img_index}: Custom Cells Images Cropped Successfully!")
        
            
        # Performing OCR on the cropped images
        extracted_text_list.extend(process_image_and_ocr(cropped_images))

        # Extracted text processing
        df = pd.DataFrame({"PDF OCR Text": extracted_text_list})
        df = df.replace('\n',' ', regex=True)

        
        for index, Filters in df.iterrows():
            # print(Filters[0])
            # Filters = Filters.iloc[0].replace("Narne", "Name").replace("Narmne", "Name").replace("Namepalli", "namepalli").replace("Father's Narme", "\nFather").replace("Husband's Narmne","\nHusband").replace("Husbands Narmne","\nHusband").replace("Hushand's Narme","\nHusband").replace("Hushand's Name","\nHusband").replace("Husband's Name","\nHusband").replace("Other's Name","\nOthers").replace("Guru's Name","\nGuru's")
            # Filters = Filters.replace("Gender", "\nGender").replace("Mather's Name", "\nMother").replace("Mother's Name", "\nMother").replace("Wife's Name", "\nWife")
            # Filters = Filters.replace("House Number", "\nHouse Number").replace("House Nurber", "\nHouse Number").replace("Hose Number", "\nHouse Number").replace("pales ", "Age ")
            # Filters = Filters.replace("Age", "\nAge").replace("Father's Name", "\nFather").replace("Fathers Name", "\nFather").replace("‘", "")
            # Filters = Filters.replace("MALE", "MALE\n").replace("Name Naik","namenaik").replace("NameNaik", "namenaik").replace("Name esa4", "")
            # Filters = Filters.replace("Name Nayak", "name nayak").replace("Namena", "namena").replace("Narne", "Name").replace(" Nae: ", "Name").replace("Narme:", "Name")
            # Filters = Filters.replace("Photo is  Available", "").replace("Photo is", "").replace("Available", "").replace("Nare:", "Name")
            # Filters = Filters.replace("Name", "\nName").replace("\x0c", "").replace(":", "").replace("=", "").replace(";","")
            # Filters = Filters.replace(",","").replace("漏","").replace(">","").replace(".","").replace("'","").replace("@", "")
            # Filters = Filters.replace("Phata is","").replace("Phato is","").replace("House Nurnber", "\nHouse Number").replace("House Nurmnber", "\nHouse Number")
            # Filters = Filters.replace("Mather's Name", "\nMother").replace("Hushand's Narme","\nHusband")

            Filters = Filters.iloc[0].replace("Narne", "Name").replace("Narme", "Name").replace(" MALESH", "malesh").replace("Narmne", "Name").replace("Namegiar", "").replace("Father's Narme", "\nFather").replace("Husband's Narmne","\nHusband").replace("Husbands Narmne","\nHusband").replace("Hushand's Narme","\nHusband").replace("Hushand's Name","\nHusband").replace("Husband's Name","\nHusband").replace("Other's Name","\nOthers").replace("Guru's Name","\nGuru's")
            Filters = Filters.replace("Gender", "\nGender").replace("Mather's Name", "\nMother").replace("Mother's Name", "\nMother").replace("Wife's Name", "\nWife")
            Filters = Filters.replace("House Number", "\nHouse Number").replace("House Nurber", "\nHouse Number").replace("Hose Number", "\nHouse Number").replace("pales ", "Age ")
            Filters = Filters.replace("Age", "\nAge").replace("Father's Name", "\nFather").replace("Fathers Name", "\nFather").replace("‘", "")
            Filters = Filters.replace("MALE", "MALE\n").replace("Name Naik","namenaik").replace("NameNaik", "namenaik").replace("Name esa4", "")
            Filters = Filters.replace("Name Nayak", "name nayak").replace("Namena", "namena").replace("Narne", "Name").replace(" Nae: ", "Name").replace("Narme:", "Name")
            Filters = Filters.replace("Photo is  Available", "").replace("Photo is", "").replace("Available", "").replace("Nare:", "Name")
            Filters = Filters.replace("Name", "\nName").replace("\x0c", "").replace(":", "").replace("=", "").replace(";","")
            Filters = Filters.replace(",","").replace("漏","").replace(">","").replace(".","").replace("'","").replace("@", "")
            Filters = Filters.replace("Phata is","").replace("Phato is","").replace("House Nurnber", "\nHouse Number").replace("House Nurmnber", "\nHouse Number")
            Filters = Filters.replace("Mather's Name", "\nMother").replace("Hushand's Narme","\nHusband")
            
            result = Filters.split('\n')

            # print(result)
            # Removing The Space Between The String From The VoterIDs
            result[0]=result[0].replace(" ","")
            result[-1]=result[-1].replace(" ","")
            
            # Extracting The VoterIds From The Raw Data And Appending It To The Dictionary
            # length of the VoterID String should greaterthan 4
            if not len(result)<=4:
                concatenated_data = '\n'.join(result)
                Voterids = re.findall(r"[a-zA-Z]{2,}\d{6,}", concatenated_data)
                if Voterids:
                    if len(Voterids[0])> 10:
                        
                        final_id=Voterids[0][-10:]
                        Data_Dict.setdefault("Voterids", []).append(final_id.strip().replace(" ",""))
                    else:
                        Data_Dict.setdefault("Voterids", []).append(Voterids[0].strip().replace(" ",""))
                else:
                    Data_Dict.setdefault("Voterids", []).append("NULL")

                
                # Setting The Flag = True to Gaurdian Name From the Data 
                relationship_types = ["Mother", "Husband", "Father", "Others", "Wife", "Guru's"]
                relation_found = False
                HNO=[]
                Age=[]
                Gender=[]
                for idx, Data in enumerate(result):
                    # print(Data)

                    if Data.startswith("Name"):
                        Data=Data.replace("!","I")
                        Data=re.sub('[^A-Za-z0-9 ]+', '', Data)
                        Data=Data.replace("Name","")
                        Data_Dict.setdefault("Name", []).append(Data.strip())

                    if Data.startswith("House Number"):
                        Data=re.sub('[^A-Za-z0-9 -/]+', '', Data)
                        Data=Data.replace("House Number","HNO").replace('"','')
                        Data=Data.lstrip('-').rstrip('-')
                        HNO.append(Data)
                        # Data_Dict.setdefault("House Number", []).append(Data.strip())

                    if Data.startswith("Age"):
                        Data=Data.strip().rstrip('-')
                        Data=re.sub('[^A-Za-z0-9 ]+', '', Data)
                        Data=Data.replace("Age","")
                        Age.append(Data)
                        # Data_Dict.setdefault("Age", []).append(Data.strip())

                    if Data.startswith("Gender"):
                        Data=re.sub('[^A-Za-z0-9 ]+', '', Data)
                        Data=Data.replace("Gender","")
                        Gender.append(Data)
                        # Data_Dict.setdefault("Gender", []).append(Data.strip())

                    for relation in relationship_types:
                        if Data.startswith(relation):
                            Data=Data.replace("-","").strip()
                            Data=re.sub('[^A-Za-z0-9 ]+', '', Data)
                            Data_Dict.setdefault("GaurdianName", []).append(Data.strip())
                            relation_found = True
                if HNO:
                    # print(HNO[-1])
                    Data_Dict.setdefault("House Number", []).append(HNO[-1].strip())
                if Age:
                    # print(Age[-1])
                    Data_Dict.setdefault("Age", []).append(Age[-1].strip())
                if Gender:
                    # print(Gender[-1])
                    Data_Dict.setdefault("Gender", []).append(Gender[-1].strip())
                
                if not relation_found:
                    Data_Dict.setdefault("GaurdianName", []).append("Unknown Relationship")

                # Reset the flag for the next iteration
                relation_found = False

        # Finding the maximum length among the lists
        max_length = max(len(lst) for lst in Data_Dict.values())

        # Pading the shorter lists with None
        for key, value in Data_Dict.items():
            Data_Dict[key] = value + [None] * (max_length - len(value))

        # Convert dictionary to DataFrame
        df = pd.DataFrame(Data_Dict)
            
        # columns to apply the default values 
        selected_columns = ["MainTown/Village", "PoliceStation", "Mandal", "RevenueDivision",
                            "District", "PinCode","AssemblyConstituency",
                            "ParliamentaryConstituency","PollingStation"]

        # iloc to select the first row for selected columns
        default_values = df[selected_columns].iloc[0]

        # Using fillna to fill missing values with the default values in selected columns
        df[selected_columns] = df[selected_columns].fillna(default_values)

        # updated DataFrame
        print(df)


        # Check if any row is None
        any_none_row = df.isna().any(axis=1).any()

        # Save the DataFrame to a CSV file in the same directory as the PDF
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        csv_file_path = os.path.join(os.path.dirname(file_path), f"{base_name}.csv")
        # df.to_csv(csv_file_path, index=False)
        # print("CSV saved to:", csv_file_path)
        save_location = r"D:\CMA\NoneData"

        # Ensure the specified directory exists, including the "NoneData" folder
        os.makedirs(save_location, exist_ok=True)

        # Save the text file with base_name inside the "Data" folder
        text_file_path = os.path.join(save_location, f"{base_name}.txt")

        if any_none_row:
            print(f"There is at least one row with a None value in PDF {base_name}.pdf")

            # Check if the file already exists
            if os.path.exists(text_file_path):
                # Remove the old file
                os.remove(text_file_path)

            with open(text_file_path, 'w') as text_file:
                text_file.write(f"There is at least one row with a None value in PDF {base_name}.pdf \n{file_path}\n")
            
            print(f"NoneData Text file saved at: {text_file_path}")
            # print(f"CSV Is Not Genereted For The PDF {file_path}")
            # Save the DataFrame to a CSV file in the same directory as the PDF

            df.to_csv(csv_file_path, index=False)
            print("NoneData CSV saved to:", csv_file_path)
        else:

            # Check if the file already exists
            if os.path.exists(text_file_path):
                # Remove the old file
                os.remove(text_file_path)
            
            # Save the DataFrame to a CSV file in the same directory as the PDF
            df.to_csv(csv_file_path, index=False)            
            print("No row contains None values.")
            print("CSV saved to:", csv_file_path)
            

        # Record end time
        end_time = datetime.datetime.now()
        print("Program Execution End Time!",end_time)

        # Calculate duration
        duration = end_time - start_time
        print(f"Duration: {duration}")

        # If the code executed without any exceptions, remove the error log file
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        error_log_file_path = os.path.join(r"D:\CMA\ErrorLogs", f"{base_name}.txt")

        if os.path.exists(error_log_file_path):
            os.remove(error_log_file_path)
            print(f"Error log file removed: {error_log_file_path}")

        return cropped_images, extracted_text_list
    
    except Exception as e:

        # Save the DataFrame to a CSV file in the same directory as the PDF
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        csv_file_path = os.path.join(os.path.dirname(file_path), f"{base_name}.csv")
        # Capture the exception and log the error along with the file path
        error_message = f"Error processing file: {file_path}\n"
        error_message += f"Error details: {str(e)}\n"

        # location to save the error log file
        save_location = r"D:\CMA\ErrorLogs"

        # Ensure the specified directory exists, including the "ErrorLogs" folder
        os.makedirs(save_location, exist_ok=True)
        
        # Save the error log file
        error_log_file_path = os.path.join(save_location, f"{base_name}.txt")

        # Check if the file already exists
        if os.path.exists(error_log_file_path):
            # Remove the old file
            os.remove(error_log_file_path)

        with open(error_log_file_path, 'a') as error_log_file:
            error_log_file.write(error_message)

        print(f"Error log saved at: {error_log_file_path}")

# Allow user to select multiple PDF files
file_paths = filedialog.askopenfilenames(title="Select PDF Files", filetypes=[("PDF files", "*.pdf")])
print(f"Number Of Selected PDF Files {len(file_paths)}")

# Set the output directory
output_directory = r'D:\CMA\Images'

# Ensure the specified directory exists
os.makedirs(output_directory, exist_ok=True)

# Number of threads in the thread pool
num_threads = 5

# Use ThreadPoolExecutor for parallel execution
with ThreadPoolExecutor(max_workers=num_threads) as executor:
    
    # Submit tasks for each PDF file
    futures = [executor.submit(process_pdf_file, file_path, output_directory) for file_path in file_paths]

    # Wait for all tasks to complete
    for future in futures:
        future.result()
    # Print statement after processing all PDF files
    print("All PDF Files Processed Successfully.")