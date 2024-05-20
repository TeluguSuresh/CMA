from tkinter import filedialog
from pdf2image import convert_from_path
from PyPDF2 import PdfReader
import os
import pytesseract
import re
import pandas as pd
import datetime
from concurrent.futures import ThreadPoolExecutor
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
        images = convert_from_path(pdf_path, first_page=3, last_page=3, fmt="jpeg", poppler_path=poppler_path)
    
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
            }
        
        # croping custom cells from the image
        for img_index, img in enumerate(images):
  
            # img_index>1 is ignoring the second Image/Page From The PDF And Performing The Other Process After The Second Page
                # Original_Size=img.size
                # img.save(output_directory,f"FullImage{img_index}.jpg")
                # # print(f"Original Size of the Image {Original_Size} , image index is {img_index}")
                # if Original_Size!=(1653, 2339):
                #     print(f"Image Size is Different{Original_Size} , image index is {img_index}")
                # else:
                #     pass
                
                cropped_image = img.crop((70, 115, 1580, 2140))
                output_path = os.path.join(output_directory, f"FullImage{img_index}.jpg")
                cropped_image.save(output_path)

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
                    cell.save(output_path)
                    cropped_images.append(cell)
                print(f"{img_index}: Custom cell Images cropped successfully!")
        
            
        # Performing OCR on the cropped images
        extracted_text_list.extend(process_image_and_ocr(cropped_images))

        # Extracted text processing
        df = pd.DataFrame({"PDF OCR Text": extracted_text_list})
        df = df.replace('\n',' ', regex=True)

        for index, Filters in df.iterrows():
            print(Filters[0])
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
        
            # print(Filters)
            result = Filters.split('\n')
            # print(result)

            # print(result)
            # Removing The Space Between The String From The VoterIDs
            result[0]=result[0].replace(" ","")
            result[-1]=result[-1].replace(" ","")
            
            # Extracting The VoterIds From The Raw Data And Appending It To The Dictionary
            if not len(result)<=4:
                concatenated_data = '\n'.join(result)


                Voterids = re.findall(r"[a-zA-Z]{2,}\d{6,}", concatenated_data)

                Voterids = re.findall(r"\d{6,}", concatenated_data)
                if Voterids:
                    # print(Voterids[0])
                    if len(Voterids[0])> 10:
                        
                        final_id=Voterids[0][-10:]
                        # print(final_id)
                        final_id="SOT"+final_id
                        Data_Dict.setdefault("Voterids", []).append(final_id.strip().replace(" ",""))
                    else:
                        # print(Voterids[0])
                        Voterids[0]="SOT"+Voterids[0]
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
                        Data=Data.replace("Name","").replace("-","")
                        Data_Dict.setdefault("Name", []).append(Data.strip())

                    if Data.startswith("House Number"):

                        Data=Data.lstrip('-').rstrip('-')
                        Data=re.sub('[^A-Za-z0-9 -/]+', '', Data)
                        Data=Data.replace("House Number","HNO").replace("漏","").replace("©","").replace('"','')
                        HNO.append(Data)

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
                print("*"*20)

                if HNO:
                    # print(HNO[-1])
                    Data_Dict.setdefault("House Number", []).append(HNO[-1].strip())
                if Age:
                    # print(Age[-1])
                    Data_Dict.setdefault("Age", []).append(Age[-1].strip())
                if Gender:
                    # print(Gender[-1])
                    Data_Dict.setdefault("Gender", []).append(Gender[-1].strip())



                # print(f"House Number Length {len(Data.strip())}")
                if not relation_found:
                    Data_Dict.setdefault("GaurdianName", []).append("Unknown Relationship")

                # Reset the flag for the next iteration
                relation_found = False
                # print("*"*30)

        # Finding the maximum length among the lists
        max_length = max(len(lst) for lst in Data_Dict.values())

        # Pading the shorter lists with None
        for key, value in Data_Dict.items():
            Data_Dict[key] = value + [None] * (max_length - len(value))

        # Convert dictionary to DataFrame
        df = pd.DataFrame(Data_Dict)
            

        # updated DataFrame
        print(df)


        # Check if any row is None
        any_none_row = df.isna().any(axis=1).any()

        # Save the DataFrame to a CSV file in the same directory as the PDF
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        csv_file_path = os.path.join(os.path.dirname(file_path), f"{base_name}.csv")
        df.to_csv(csv_file_path, index=False)
        print("CSV saved to:", csv_file_path)

        if any_none_row:
            print(f"There is at least one row with a None value in PDF {base_name}.pdf","*"*50)

            save_location = r"D:\CMA\NoneData"

            # Ensure the specified directory exists, including the "NoneData" folder
            os.makedirs(save_location, exist_ok=True)

            # Save the text file with base_name inside the "Data" folder
            text_file_path = os.path.join(save_location, f"{base_name}.txt")

            with open(text_file_path, 'w') as text_file:
                text_file.write(f"There is at least one row with a None value in PDF {base_name}.pdf \n{file_path}\n And CSV Is Not generated For this FDF {base_name}")
            
            print(f"Text file saved at: {text_file_path}")

        else:        
            print("No row contains None values.")

            

        # Record end time
        end_time = datetime.datetime.now()
        print("Program Execution End Time!",end_time)

        # Calculate duration
        duration = end_time - start_time
        print(f"Duration: {duration}")

        return cropped_images, extracted_text_list
    

# Allow user to select multiple PDF files
file_paths = filedialog.askopenfilenames(title="Select PDF Files", filetypes=[("PDF files", "*.pdf")])

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