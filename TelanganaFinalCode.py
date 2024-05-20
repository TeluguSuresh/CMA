from tkinter import filedialog
from pdf2image import convert_from_path
from PyPDF2 import PdfReader
import os
import pytesseract
import re
import pandas as pd
import datetime
from concurrent.futures import ThreadPoolExecutor
import cv2
import numpy as np

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


def preprocess_image(img):
    # Convert image to grayscale
    gray = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2GRAY)
    
    # Apply Gaussian blur
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    
    # Apply adaptive thresholding
    threshold_img = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 4)
    
    return threshold_img


def process_image_and_ocr(cropped_images):

    extracted_text = []

    for i, CropedImage in enumerate(cropped_images):
        # Performing OCR on the cropped image
        # text = pytesseract.image_to_string(CropedImage, lang='eng')

        # Inside the loop where you process each cropped image
        preprocessed_image = preprocess_image(CropedImage)

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
            "PostOffice":[],
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
                output_path = os.path.join(output_directory, f"FullImage{img_index}.jpg")
                    
                # croping custom cells from the first image
                cropped_image = img.crop((30, 185, 1625, 2110))
                # cropped_image.save(output_path)

                # Coordinates for custom cells
                custom_cells = [
                    (0, 0, 1290, 75),
                    (0, 85, 1590, 150),
                    (645, 600, 1500, 1150),
                    (0, 1210, 820, 1650),
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
                                        
                    i=re.sub('[^A-Za-z0-9 \n,()]+', '', i)
                    
                    # SplitData=i.split("\n")
                    # print(indexes,SplitData)
                    if indexes==0:
                        print("Index is 0")
                        # print(i)
                        i=i.replace("Status of Assembly Constituency","\nAssembly Constituency")

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
                        i=i.replace('\n','').replace("Assembly Constituency is located","\nParliamentary Constituency")
                        SplitData=i.split("\n")
                        for strings in SplitData:
                         if strings.startswith("Parliamentary Constituency"):
                                    strings=re.sub('[^A-Za-z0-9 ()-]+', '', strings)
                                    print(strings)
                                    strings=strings.replace("Parliamentary Constituency","")
                                    Data_Dict["ParliamentaryConstituency"].append(strings.strip())
                    
                    if indexes==2:
                        print("Index is 2")
                        i=i.replace("\x0c","")

                        i=i.replace("\n"," ")
                        i=i.replace("Main Town","\nMain Town").replace("Post Office","\nPost Office").replace("Police Station","\nPolice Station").replace("Mandal","\nMandal")
                        i=i.replace("Revenue Division","\nRevenue Division").replace("District","\nDistrict").replace("Pin code","\nPin Code").strip()
                        
                        SplitData=i.split("\n")

                        for strings in SplitData:

                            if strings.startswith("Main Town"):

                                    print(strings)
                                    strings=strings.replace("Main Town or Village","")
                                    Data_Dict["MainTown/Village"].append(strings.strip())

                            if strings.startswith("Post Office"):
                                print(strings)
                                strings=strings.replace("Post Office","")
                                Data_Dict["PostOffice"].append(strings.strip())


                            if strings.startswith("Police Station"):
                                    print(strings)
                                    strings=strings.replace("Police Station","")
                                    Data_Dict["PoliceStation"].append(strings.strip())

                            if strings.startswith("Mandal"):
                                    print(strings)
                                    strings=strings.replace("Mandal","")
                                    Data_Dict["Mandal"].append(strings.strip())

                            if strings.startswith("Revenue Division"):
                                    print(strings)
                                    strings=strings.replace("Revenue Division","")
                                    Data_Dict["RevenueDivision"].append(strings.strip())

                            if strings.startswith("District"):
                                    print(strings)
                                    strings=strings.replace("District","")
                                    Data_Dict["District"].append(strings.strip())

                            if strings.startswith("Pin Code"):
                                    print(strings)
                                    strings=strings.replace("Pin Code","")
                                    Data_Dict["PinCode"].append(strings.strip())

                    if indexes==3:
                        print("Index is 3")
                        i=i.replace("\n","").replace("No and Name of Polling Station","\nName of Polling Station")
                        i=i.replace("Address of Polling Station","\nAddress of Polling Station")
                        # print(i)

                        SplitData=i.split("\n")    
                        for Polling in SplitData:
                            # print(Polling)
                            if Polling.startswith("Name of Polling Station"):
                                print(Polling)
                                Polling=Polling.replace("Name of Polling Station","").strip()
                                Data_Dict["PollingStation"].append(Polling)
                  
                
            # img_index>1 is ignoring the second Image/Page From The PDF And Performing The Other Process After The Second Page
            if img_index>1:
                cropped_image = img.crop((35, 75, 1620, 2280))
                output_path = os.path.join(output_directory, f"FullImage{img_index}.jpg")
                # cropped_image.save(output_path)

                custom_cells = [
                    (0, 0, 520, 220),     (520, 0, 1055, 220),     (1055, 0, 1580, 220),
                    (0, 210, 520, 440),   (520, 210, 1055, 440),   (1055, 210, 1580, 440),
                    (0, 420, 520, 660),   (520, 420, 1055, 660),   (1055, 420, 1580, 660),
                    (0, 640, 520, 880),   (520, 640, 1055, 880),   (1055, 640, 1580, 880),
                    (0, 880, 520, 1100),  (520, 880, 1055, 1100),  (1055, 880, 1580, 1100),
                    (0, 1100, 520, 1320),  (520, 1100, 1055, 1320),  (1055, 1100, 1580, 1320),
                    (0, 1310, 520, 1540), (520, 1310, 1055, 1540), (1055, 1310, 1580, 1540),
                    (0, 1520, 520, 1760), (520, 1520, 1055, 1760), (1055, 1520, 1580, 1760),
                    (0, 1740, 520, 1980), (520, 1740, 1055, 1980), (1055, 1740, 1580, 1980),
                    (0, 1980, 520, 2200), (520, 1980, 1055, 2200), (1055, 1980, 1580, 2200),
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
            Filters = Filters.iloc[0]
            Filters=re.sub('[^A-Za-z0-9 \n,()]+','', Filters)

            # if index==1:
            #     break
            Filters = Filters.replace("Narmne","Name").replace("Narne","Name").replace("Fathers Name","\nFather").replace("Fathers Narme","\nFather").replace("Husbands Narmne","\nHusband").replace("Husbands Narmne","\nHusband").replace("Hushands Narme","\nHusband")
            Filters = Filters.replace("Hushands Name","\nHusband").replace("Husbands Name","\nHusband").replace("Others Name","\nOthers").replace("Gurus Name","\nGurus").replace("Others","\nOthers")
            Filters = Filters.replace("Gender","\nGender").replace("Gander","\nGender").replace("Gonder","\nGender").replace("Gendar","\nGender").replace("Mathers Name","\nMother").replace("Mothers Name","\nMother").replace("Wifes Name","\nWife")
            Filters = Filters.replace("House","\nHouse Number").replace("Ago ","Age ").replace("Gonder  F","Gender F").replace("Gonder  M","Gender M").replace("Husbands Narme", "\nHusband")
            Filters = Filters.replace("Age", "\nAge").replace("Fathers Name","\nFather").replace("Fathors","Fathers").replace("Fathars","Fathers").replace("Husbands Namne","\nHusband")
            Filters = Filters.replace("Fathers Namne", "\nFather").replace("Fathers Namo", "\nFather").replace("Fathers Namne", "\nFather").replace("Fathers Nano", "\nFather")
            Filters = Filters.replace("Name Naik","namenaik").replace("NameNaik","namenaik").replace("Name esa4", "").replace("Fathers Nane","\nFather").replace("Mathers Narn","\nMother")
            Filters = Filters.replace("Name Nayak","name nayak").replace("Namena","namena").replace("Narne","Name")
            Filters = Filters.replace("Photo is  Available", "").replace("Photo is","").replace("Available","").replace("Avallable","")
            Filters = Filters.replace("Name", "\nName").replace("Photo","").replace("Proto","").replace("Pnoto","").replace("photo","")
            Filters = Filters.replace("Phata is","").replace("Phato is","").replace("House Nurnber","\nHouse Number").replace("House Nurmnber","\nHouse Number")
            Filters = Filters.replace("Mathers Name","\nMother").replace("Hushands Narme","\nHusband")
            
            # print(Filters)
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
                        Data=Data.replace("Name","")
                        Data_Dict.setdefault("Name", []).append(Data.strip())

                    if Data.startswith("House Number"):
                        Data=Data.replace("House Number","HNO").replace('Number','')
                        Data=Data.lstrip('-').rstrip('-')
                        HNO.append(Data)
                        # Data_Dict.setdefault("House Number", []).append(Data.strip())

                    if Data.startswith("Age"):
                        Data=Data.strip().rstrip('-')
                        Data=Data.replace("Age","")
                        Age.append(Data)
                        # Data_Dict.setdefault("Age", []).append(Data.strip())

                    if Data.startswith("Gender"):
                        Data=Data.replace("Gender","")
                        Gender.append(Data)
                        # Data_Dict.setdefault("Gender", []).append(Data.strip())

                    for relation in relationship_types:
                        if Data.startswith(relation):
                            Data_Dict.setdefault("GaurdianName", []).append(Data.strip())
                            relation_found = True
                if HNO:
                    # print(HNO[-1])
                    Data_Dict.setdefault("House Number", []).append(HNO[-1].strip())
                
                if len(Age)==0:
                    Data_Dict.setdefault("Age", []).append("Null")
                else:
                    # print(Age[-1])
                    Data_Dict.setdefault("Age", []).append(Age[-1].strip())

                if len(Gender)==0:
                    # print("Gender is Null")
                    Data_Dict.setdefault("Gender", []).append("Null")
                else:
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
        selected_columns = ["MainTown/Village","PostOffice", "PoliceStation", "Mandal", "RevenueDivision",
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
        NoneData_location=os.path.dirname(file_path)
        # df.to_csv(csv_file_path, index=False)
        # print("CSV saved to:", csv_file_path)

        #None Data Folder creation based on the selected filepath
        save_location = f"{NoneData_location}/NoneData"

        # saving NoneData File in the NoneData Folder
        NoneData_csv_file_path = os.path.join(save_location, f"{base_name}.csv")
        # save_location= r""

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
            # None Data File is Saving at the same directory as textfile is in
            df.to_csv(NoneData_csv_file_path, index=False)
            print("NoneData CSV saved to:", NoneData_csv_file_path)
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