from PyPDF2 import PdfFileMerger, PdfReader, PdfWriter
import os
import schedule
import time
import datetime
import shutil
import logging

def getListOfFiles(dirName):
    # ==========================================
    # create a list of file and sub directories
    # ==========================================
    # names in the given directory 
    listOfFile = os.listdir(dirName)
    allFiles = list()
    # Iterate over all the entries
    for entry in listOfFile:
        # Create full path
        fullPath = os.path.join(dirName, entry)
        # If entry is a directory then get the list of files in this directory 
        if os.path.isdir(fullPath):
            allFiles = allFiles + getListOfFiles(fullPath)
        else:
            allFiles.append(fullPath)
                
    return allFiles
    
def AddMetadata(pdfFileInput,pdfFileOutput):
        
    # Add the metadata
    input_info=pdfFileInput.getDocumentInfo()
    # print(input_info)
    #Author
    if hasattr(input_info, "author"):
        output_Author=input_info.author
    else:
        output_Author="Pierrick BLAISE"
    #Title
    if hasattr(input_info, "title"):
        output_Title=input_info.title
    else:
        output_Title="Quotation Order Invoice"
    #Subject
    if hasattr(input_info, "subject"):
        output_Subject=input_info.subject
    else:
        output_Subject="Document from SaintLu Metal"
    #ModDate
    now = datetime.datetime.now()
    output_ModDate=now.strftime("D:%Y%m%d%H%M%SZ")
    #CreationDate 'D:20210831155332Z',
    try:
        output_CreationDate = pdfFileInput.documentInfo['/CreationDate']
    except:
        output_CreationDate = output_ModDate
    pdfFileOutput.addMetadata(
        {
            "/Author": str(output_Author),
            "/Producer": "PDFLogo for SaintLu by Pierrick BLAISE",
            "/CreationDate": str(output_CreationDate),
            "/ModDate": str(output_ModDate),
            "/Title": str(output_Title),
            "/Subject": str(output_Subject),
            "/Creator": "pyPDF2"
        }
    )
    print(datetime.datetime.now().isoformat())
    
def AddSaintLuLogo(filePath):
    for pdf_file in getListOfFiles(filePath):
        # List the file ending with _NoLogo.pdf
        if pdf_file.endswith("_NoLogo.pdf"):
            try:
                NoLogo_file = pdf_file
                #Display the list of files
                #print(os.path.abspath(NoLogo_file))
                NoLogo_filepath = os.path.abspath(NoLogo_file)
                logging.info('New file detected %s',NoLogo_filepath)
                if pdf_file.endswith("_100_NoLogo.pdf"):
                    WithLogo_filepath = NoLogo_filepath.replace('_100_NoLogo','_Original')
                elif pdf_file.endswith("_010_NoLogo.pdf"):
                    WithLogo_filepath = NoLogo_filepath.replace('_010_NoLogo','_Copy')
                elif pdf_file.endswith("_001_NoLogo.pdf"):
                    WithLogo_filepath = NoLogo_filepath.replace('_001_NoLogo','_Trial')
                else:
                    WithLogo_filepath = NoLogo_filepath.replace('_NoLogo','')
                
                
                #Display the list of files that will be created
                #print(WithLogo_filepath)
                print('---------------------- New File Detected --------------------------')
                print('Adding the Logo on file ',os.path.basename(WithLogo_filepath))
                
                # Link to the PDF With the logo that will be applied
                logo = r'\\SERVER1\ServerData\Projects\_AddLogoToPdf\LetterHead.DONOTMOVE.pdf'
                

                # Test if the file with logo already exist
                if os.path.exists(WithLogo_filepath):
                    logging.info('The file %s already exist. Trying again in 2s',os.path.basename(WithLogo_filepath))
                    time.sleep(2)
                    if os.path.exists(WithLogo_filepath):
                        logging.info('The file %s already exist',os.path.basename(WithLogo_filepath))
                        logging.info('Adding a Timestamp at the end of the file')
                        now = datetime.datetime.now()
                        date_time = now.strftime("%Y%m%d%H%M%S")
                        WithLogo_filepath = WithLogo_filepath.replace('.pdf','_')
                        WithLogo_filepath = WithLogo_filepath + date_time + str(".pdf")
                
                # Name of the file that will be created
                WithLogo = WithLogo_filepath
                
                # Open the file with logo
                WithLogo_file = open(WithLogo, "wb")
                logging.info('Opening final file %s',os.path.basename(WithLogo_filepath))
                # Open the NoLogo File, Merge it with the Logo and save it
                with open(NoLogo_file, "rb") as input_file, open(logo, "rb") as logo_file:
                    logging.info('Opening nologo file %s',os.path.basename(NoLogo_filepath))
                    input_pdf = PdfReader(input_file)
                    logo_pdf = PdfReader(logo_file)
                    logo_page = logo_pdf.pages[0]

                    output = PdfWriter()

                    for i in range(len(input_pdf.pages)):
                        pdf_page = input_pdf.pages[i]
                        if i == 0:
                            pdf_page.merge_page(logo_page)
                        output.add_page(pdf_page)

                    #AddMetadata
                    # Need to update to pyPDF2 3.0.0
                    # AddMetadata(input_pdf,output)
                    
                    output.write(WithLogo_file)
                    logging.info('Writing in final file')
                    
                    
                WithLogo_file.close()
                logo_file.close()
                input_file.close()
                #Delete the "NoLogo" File
                #os.remove(NoLogo_filepath)

                # Move the noLogo File to Archives
                try:
                    newpath = os.path.join(r'\\SERVER1\ServerData\Archives',os.path.basename(NoLogo_filepath))
                    if not os.path.exists(r'\\SERVER1\ServerData\Archives'):
                        os.makedirs(r'\\SERVER1\ServerData\Archives')
                    shutil.move(NoLogo_filepath,newpath)
                    
                    print('Moving the file to Archives',os.path.basename(NoLogo_filepath))
                    logging.info('File Moved to %s', os.path.abspath(NoLogo_filepath))
                    print('--------------------------- Success -------------------------------')
                    # Save the new PDF as a file
                    
                except Exception as e:
                    pass
                    # If the file cannot be moved, remove the created file to avoid future conflict
                    os.remove(WithLogo_filepath)
                    print("something went wrong while moving the file:")
                    logging.error('Could not move the file to Archives')
                    print(e)
                logging.info('----------------------------------')
            except Exception as e:
                pass
                print(e)


    print('Script PDFLogo Successfully Executed at',datetime.datetime.now())
    print('Folder: ',filePath)

  
    

print('=========================== Saint Lu Metals PDFLogo =================================')
print('This script automatically add the SaintLu Logo to the PDF Files generated from Klaes.')
print('If there is any issues with this script, contact Pierrick Blaise: bpierrick@gmail.com')
print('====================== /!\ DO NOT CLOSE THIS WINDOW /!\ =============================')
#basic logging config
logFileName='AddLogoToPdf.log'
logging.basicConfig(
    filename=logFileName,
    filemode='a', 
    format='%(asctime)s - %(levelname)s:%(message)s',
    level=logging.INFO
)
logging.info('===== Script Initiated =====')

schedule.every(2).seconds.do(AddSaintLuLogo,r'\\SERVER1\ServerData\Projects')
schedule.every(15).seconds.do(AddSaintLuLogo,r'\\SERVER1\ServerData\Accounting\Klaes PriceList')
while True:
    try:
        schedule.run_pending()
    except Exception as e:
        pass
        print("something went wrong")
        print(e)
    time.sleep(1)
