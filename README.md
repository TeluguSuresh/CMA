Welcome to the CMA-Comman-Man-Army wiki! This project aims to develop a system for extracting data from scanned PDF files, which are in the form of images, and transforming the extracted data into CSV format. The CSV files are then stored in a PostgreSQL database. The process involves the use of Optical Character Recognition (OCR) to convert images to text, and Python with Pandas for data manipulation. Additionally, multithreading is used to process multiple PDFs in parallel, improving efficiency and speed.

The primary goal of this project is to automate the extraction of data from scanned PDF documents. These documents are converted into images, processed using OCR to extract textual data, which is then structured and stored in a CSV file format. Finally, the CSV files are uploaded into a PostgreSQL database for further analysis and querying. To improve processing efficiency, multithreading is utilized to handle multiple PDFs simultaneously.

Softwares Python 3.x PostgreSQL Tesseract OCR Python Libraries: pytesseract pdf2image pandas psycopg2 numpy concurrent.futures

System Architecture PDF to Image Conversion: Convert scanned PDF pages into images using pdf2image. OCR Processing: Use Tesseract OCR to extract text from images. Data Structuring: Organize the extracted text into a structured format using Pandas DataFrame. CSV Generation: Export the structured data into CSV files. Database Storage: Insert the CSV data into a PostgreSQL database. Parallel Processing: Use multithreading to process multiple PDFs concurrently.

Prerequisites Install Python Install PostgreSQL Install Tesseract OCR

End-to-End Testing Performed end-to-end testing by running the entire workflow on a sample PDF and verifying the results in the PostgreSQL database.
