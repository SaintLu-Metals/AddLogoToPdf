from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
import os
import schedule
import time
import datetime
import shutil

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
        if pdf_file.endswith("_NoLogo1.pdf"):
            NoLogo_file = pdf_file
            #Display the list of files
            #print(os.path.abspath(NoLogo_file))
            NoLogo_filepath = os.path.abspath(NoLogo_file)
            if pdf_file.endswith("_100_NoLogo1.pdf"):
                WithLogo_filepath = NoLogo_filepath.replace('_100_NoLogo1','_Original')
            elif pdf_file.endswith("_010_NoLogo1.pdf"):
                WithLogo_filepath = NoLogo_filepath.replace('_010_NoLogo1','_Copy')
            elif pdf_file.endswith("_001_NoLogo1.pdf"):
                WithLogo_filepath = NoLogo_filepath.replace('_001_NoLogo1','_Trial')
            else:
                WithLogo_filepath = NoLogo_filepath.replace('_NoLogo','')
            #Display the list of files that will be created
            #print(WithLogo_filepath)
            print('---------------------- New File Detected --------------------------')
            print('Adding the Logo on file ',os.path.basename(WithLogo_filepath))

            # Link to the PDF With the logo that will be applied
            logo = r'\\SERVER1\ServerData\Projects\LetterHead.DONOTMOVE.pdf'
            # Name of the file that will be created
            WithLogo = WithLogo_filepath

            # Open the NoLogo File, Merge it with the Logo and save it
            with open(NoLogo_file, "rb") as input_file, open(logo, "rb") as logo_file:
                input_pdf = PdfFileReader(input_file)
                logo_pdf = PdfFileReader(logo_file)
                logo_page = logo_pdf.getPage(0)

                output = PdfFileWriter()

                for i in range(input_pdf.getNumPages()):
                    pdf_page = input_pdf.getPage(i)
                    if i == 0:
                        pdf_page.mergePage(logo_page)
                    output.addPage(pdf_page)

                #AddMetadata
                # AddMetadata(input_pdf,output)
                
                # Save the new PDF as a file
                with open(WithLogo, "wb") as WithLogo_file:
                    output.write(WithLogo_file)
                
                
            WithLogo_file.close()
            logo_file.close()
            input_file.close()
            #Delete the "NoLogo" File
            # os.remove(NoLogo_filepath)

            # Move the noLogo File to Archives

            newpath = os.path.join(r'\\SERVER1\ServerData\Archives',os.path.basename(NoLogo_filepath))
            if not os.path.exists(r'\\SERVER1\ServerData\Archives'):
                os.makedirs(r'\\SERVER1\ServerData\Archives')
            shutil.move(NoLogo_filepath,newpath)
            print('Moving the file to Archives',os.path.basename(NoLogo_filepath))
            print('--------------------------- Success -------------------------------')
    print('Script PDFLogo Successfully Executed at',datetime.datetime.now())
    print('Folder: ',filePath)

    

print('=========================== Saint Lu Metals PDFLogo =================================')
print('This script automatically add the SaintLu Logo to the PDF Files generated from Klaes.')
print('If there is any issues with this script, contact Pierrick Blaise: bpierrick@gmail.com')
print('====================== /!\ DO NOT CLOSE THIS WINDOW /!\ =============================')
schedule.every(5).seconds.do(AddSaintLuLogo,r'\\SERVER1\ServerData\Projects')
schedule.every(15).seconds.do(AddSaintLuLogo,r'\\SERVER1\ServerData\Accounting\Klaes PriceList')
while True:
    try:
        schedule.run_pending()
    except Exception as e:
        pass
        print("something went wrong")
        print(e)
    time.sleep(1)
