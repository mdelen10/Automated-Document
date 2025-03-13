#libraries
import os
import sys
import traceback
import argparse
import tkinter as tk
from tkinter import filedialog 
import pandas as pd
from docx.opc.exceptions import PackageNotFoundError
from docx import Document
from docxtpl import DocxTemplate

image_extensions = ('.png', '.jpg', '.jpeg')

def replace_screens(useCase, imagesFolder, screenToAltTextTable):
    # 1) read the Excel file
    try:
        df = pd.read_excel(screenToAltTextTable, header=None)
    except Exception as e:
        print(f"Error: Cannot read Excel file '{screenToAltTextTable}'\n\033[1mPlease make sure the Excel File is not open in another program\033[0m")
        return False

    # 2) validate excel mapping file 
    if df.empty or len(df.columns) != 2:            #  ensure we have 2 columns
        print("Error: Excel file must have only 2 columns")
        return False
    if df.iloc[0, 0] == "Screen Name" or df.iloc[0, 0] == "Image Name":   # remove column header if it exists
        df = df.iloc[1:, :]

    # 3a) ensure image folder exists
    try:    
        os.chdir(imagesFolder)
    except Exception as e:
        print(f"Error: Cannot find image folder '{imagesFolder}'")  
        return False
    # 3b) ensure each image exists
    ImagePaths = []
    altTexts = []
    imageFolderError = False
    for index, row in df.iterrows():
        curScreenName = str(df.iat[index, 0])
        image_count = -1
        if os.path.isdir(curScreenName):
            image_count = 0
            for filename in os.listdir(curScreenName):
                if filename.lower().endswith(image_extensions):
                    curScreenPath = curScreenName + "/" + filename
                    ImagePaths.append(curScreenPath)
                    curAltText = df.iat[index, 1]
                    altTexts.append(curAltText)
                    #print(f"DEBUG  added {curScreenPath}  + {curAltText}")    
                    image_count += 1
        if image_count == -1:
            imageFolderError = True
            print(f"\033[1mError: Image folder '{curScreenName}' in excel row {index+1} is not an image directory")
        if image_count == 0:
            imageFolderError = True
            print(f"\033[1mError: Image folder '{curScreenName}' in excel row {index+1} does not contain any images\033[0m")
        if image_count > 1:
            imageFolderError = True
            print(f"\033[1mError: Image folder '{curScreenName}' in excel row {index+1} contains more than 1 image\033[0m")

    # 4) replace all the Images
    if not imageFolderError:
        context = {}
        try:
            useCasePath = os.path.abspath(useCase) 
            tpl = DocxTemplate(useCasePath)
            for i in range(len(ImagePaths)):
                tpl.replace_pic(altTexts[i], ImagePaths[i])
                tpl.render(context)
        except PackageNotFoundError as p:
            print(f"Error opening Use Case file '{useCase}'\n\033[1mPlease make sure the use case Word Document is not open.\033[0m")
            return False
        except Exception as e:
            print(str(e))
            return False
        try:
            useCasePathParent = os.path.abspath(os.path.dirname(useCase)) # os.path.dirname(useCasePath) )
            cwd = os.getcwd()
            if(cwd != useCasePathParent): # nvaigate to the parent folder of the use case
                os.chdir(useCasePathParent)
            absolutePathToUseCase, nameOfNewFile = os.path.split(useCasePath)
            finalFilePath = os.path.join("Updated_Use_Case", nameOfNewFile) # save modified use case under Updated_Use_Case/<Use Case>
            if not os.path.isdir("Updated_Use_Case"):  # if updatedUseCase/ does not exist, make it
                os.mkdir("Updated_Use_Case")
            tpl.save(rf'{finalFilePath}')
        except PermissionError as p:
             print(f"\033[1mPlease make sure the output file {finalFilePath} is not open in another program\033[0m")
             return False
        except ValueError as v:
            print(f"\033[1mError finding alt text value: {str(v)}")
            print(f"Please make sure there are no newlines in your Use Case's alt texts!\033[0m")
            return False
        except Exception as e:
            print(f"\033[1mError saving Use Case file:\033[0m {str(e)}")
            return False
        print(f"\033[1mNew images used:    \033[0m" + str(ImagePaths[:2]) + " ..")
        print(f"\033[1mView new file at:  \033[0m {finalFilePath}")
    else: # there is imageFolderError 
        print(f"\033[1mPlease make sure '{imagesFolder}' is a valid folder with images\033[0m")

    return not imageFolderError

def main():
    ''' 
        Main function to validate pythons params
        argument 'u' - Use Case to be updated
        argument 'i' - Mmage folder to look into
        argument 'm' - Mapping excel file with Screen Names & Alt Texts 
    '''
    # 1) try getting inputs from command line argument
    parser = argparse.ArgumentParser(description="Update screens in Use Case documents")
    parser.add_argument("-u", "--useCase", help="Path to the Use Case Word document")
    parser.add_argument("-i", "--images",  help="Path to the Image folders")
    parser.add_argument("-m", "--mapping", help="Path to Excel with Mapping for Screen Name to Alt Texts")
    args = parser.parse_args()

    # 2) try getting inputs from dialog box 
    if not args.useCase or not os.path.isfile(args.useCase) or len(args.useCase) < 5 or args.useCase[-5:] != ".docx":
        args.useCase = filedialog.askopenfilename(title = "Please select the Word File", filetypes=[("Word files", "*.docx")])
    if not args.images or not os.path.isdir(args.images):
        args.images = filedialog.askdirectory(title = "Please select the Image Folder")
    if not args.mapping or not os.path.isfile(args.mapping) or len(args.mapping) < 5 or args.mapping[-5:] != ".xlsx":
        args.mapping = filedialog.askopenfilename(title = "Please select the Excel File \t(1st column with Image Names and 2nd column with Alt Texts)", filetypes=[("Excel files", "*.xlsx")])

    # 3) validate inputs
    inputsValid = True
    print("\n\n")
    inputs = [ ["Use Case", args.useCase] , ["Image Folder", args.images], ["Image Mapping Excel", args.mapping] ]
    for inputName, inputValue in inputs:
        if inputValue:
            print(f" + {inputName}: '{inputValue}'")
        else:
            print(f" - Error: '{inputName}' is invalid")
            inputsValid = False
    
    success = False
    if inputsValid: 
        success = replace_screens(args.useCase, args.images, args.mapping)
        
    if  success:
        print("\x1b[6;30;42m" + "\n-------------------------\nScreen update success!\n-------------------------" + "\x1b[0m ")
    else:
        print("\x1b[6;30;41m" + "\n-------------------------\nScreen update fail!\n-------------------------" + "\x1b[0m")
    restartBool = input('\033[1mWould you like to run the program again? (y/n)\033[0m ')
    if restartBool.lower().lstrip().rstrip() == 'y':
        args.useCase = None
        args.images = None  
        args.mapping = None
        main()


if __name__ == "__main__":
    main()
