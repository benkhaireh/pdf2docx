import os, shutil
from pdf2docx import parse
import time
from shutil import copyfile
from pathlib import Path

class bcolors:
    ENDC = '\033[0m'
    PRIMARYBOLD = '\033[1;34m'
    INFOBOLD = '\033[1;36m'
    SUCCESSBOLD = '\033[1;32m'
    DANGERBOLD = '\033[1;31m'
    WARNINGBOLD = '\033[1;93m'
## Presentation
print('\n', bcolors.WARNINGBOLD+"Hi, my name is Ropyto and i'm going to help you converting all your pdf files to Word (.docx)"+bcolors.ENDC, '\n')
## --------------------------------------------
## Requesting target folder
setPath = input(bcolors.PRIMARYBOLD+"So please enter the target folder: "+bcolors.ENDC)
## Checking if the target is not empty
if not os.listdir(setPath):
    print('\n', bcolors.DANGERBOLD+'Sorry, but '+bcolors.INFOBOLD+'./'+setPath+'/'+bcolors.ENDC+bcolors.DANGERBOLD+' is empty! Please put some files inside it.'+bcolors.ENDC, '\n')
    exit()
path = setPath
## --------------------------------------------
## Requesting output folder
setTarget = input(bcolors.PRIMARYBOLD+"And enter the output folder: "+bcolors.ENDC)
## Creating tracker folder to arrange pdf files and to esaly convert
if not os.path.isdir('./tracker'):
    os.mkdir('./tracker')
tracker_dir = './tracker'
if not os.path.isdir(tracker_dir):
    os.mkdir(tracker_dir)

## Converting process
## Finding all PDF files in the target
print('\n', bcolors.WARNINGBOLD+"I'm going to find all pdf files in your targert now"+bcolors.ENDC, '\n')
counter = 0
tic = time.perf_counter()
for root, dirs, files in os.walk(path):
    for file in files:
        if file.endswith(".pdf"):
            filepath = os.path.join(root, file)
            copyfile(filepath, os.path.join(tracker_dir, file))
            counter += 1
            print('\n', bcolors.WARNINGBOLD+str(counter)+' found'+bcolors.ENDC, '\n')
toc = time.perf_counter()
print('\n', bcolors.SUCCESSBOLD+'Done '+str(counter)+f" PDF file(s) found within {toc - tic:0.2f} seconds"+bcolors.ENDC, '\n')

## Converting all PDF files to WORD (DOCX)
print('\n', bcolors.WARNINGBOLD+"I'm going to convert all now"+bcolors.ENDC, '\n')
counter = 0
tic = time.perf_counter()
for file in os.listdir(tracker_dir):
    inputFile = os.path.join(tracker_dir, file)
    filename, file_extension = os.path.splitext(file)
    docfile = filename+'.docx'
    outputFile = os.path.join(setTarget, docfile)
    parse(inputFile, outputFile)
    convertedFile = Path(outputFile)
    if not convertedFile.exists():
        print('\n', bcolors.DANGERBOLD+'Sorry, but '+bcolors.INFOBOLD+filename+'.pdf'+bcolors.ENDC+bcolors.DANGERBOLD+' can\'t be converted.'+bcolors.ENDC, '\n')
        exit()
    counter += 1
    print('\n', bcolors.WARNINGBOLD+str(counter)+' converted until now'+bcolors.ENDC, '\n')
toc = time.perf_counter()

print('\n', bcolors.SUCCESSBOLD, "Done I finish converting "+str(counter)+f" PDF file(s) to Word within {toc - tic:0.2f} seconds", bcolors.ENDC, '\n')
## Requesting to delete the tracker folder
requestForDelete = input(bcolors.INFOBOLD+"I have created a copy of all pdf files in a folder named \"tracker\" !\n So do you wanna keep it ? Answer with y(yes) or n(no): "+bcolors.ENDC)

if requestForDelete == 'n':
    for filename in os.listdir(tracker_dir):
        file_path = os.path.join(tracker_dir, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))
            break
    shutil.rmtree(tracker_dir)
    print('\n', bcolors.SUCCESSBOLD, "Done all was delete", bcolors.ENDC, '\n')
else:
    print('\n', bcolors.SUCCESSBOLD, "Good i'm not going to delete it", bcolors.ENDC, '\n')