import os
import pandas as pd
import re
import PyPDF2 as pdfReader
from PIL import Image
from win32com.client import Dispatch
import zipfile

class Document:

    """
    INIT METHOD
    """

    def __init__(self, rootDirPath, fullPath, indexNum=0, indexDigitsSize = 5, maxSizeForProcessing = 0, ignoreMaxPathLenFlag = False):

        """
        CONSTRUCTOR ARGS
        
        fullPath - MANDATORY -
        The full file path of the target file.

        index -
        An arbitrary number indicating the index position of this Document relative to something else (e.g. a list of Documents). Default is 0.

        indexDigitSize -
        The number of zeroes to prefix the index with. Default is 5, so index will be '####X' where X is the index and # is zeros if not used.

        maxSizeForProcessing -
        The maximum size of file in bytes to be processed. Default is 0, which means no limit.
        This is helpful if you want to avoid processing very large files.
        Setting this to 1 is also helpful if you only want to process the file size of files and nothing else.

        ignoreMaxPathLenFlag -
        Choose whether or not to ignore max file path length flags. Default is False.
        If your MAX_PATH variable has been modified and Python can read greater than the normal limit, you can ignore the max path flag using this argument.
        Do not use this if this is not the case, as almost all processing is impossible if the OS path limit exists.
        Normally files of greater than the Windows OS path limit are not processed. 
        """

        # Sets the index number of this document to the given param in the constructor. The default is 0.
        # Also takes an index digit size param to prefill the index with zeroes. The default is 5 digit long number.
        
        self.docIndexDigitsSize = indexDigitsSize
        self.docIndexNum = indexNum
        self.docIndexDigits = str(indexNum).zfill(indexDigitsSize)

        # Set maximum file size threshold for the more involved file processing functions. Useful if some files are ludicrously large.
        # The default is 0, meaning no limit.
        # Set the red flag indicating if the max size threshold has been exceeded is defaulted to False.
        # Set whether or not the max length flag should be ignored from the constructor argument.
        self.docMaxSizeForProcessing = maxSizeForProcessing
        self.docMaxSizeForProcessingSkipFlag = False
        self.docIgnoreMaxPathLenFlag = ignoreMaxPathLenFlag

        # Create full path and path related vars, then set these.
        # Note the full path and path only length flags (using Windows limits).
        # A long file path will outright prevent most meaningful processing of the file, therefore if this flag is on, it will disable most processing in this Document class.
        self.docRootDirPath = rootDirPath
        self.docPathFull = fullPath
        self.docPathNoRoot = ""
        self.docPathFullLen = 0
        self.docPathFullLenFlag = False
        self.docPathOnly = ""
        self.docPathOnlyLen = 0
        self.docPathOnlyLenFlag = False
        self.docPathCommaFlag = False
        self.setDocPathDetails()

        """
        TO DO - Create root path finder and add a column / var that gets the path without the common path - useful for creating the index.
        """

        # Set up Document metadata variables then begin processing.
        self.docSize = 0
        self.docSizeMsg = "Not processed"
        self.docNumPages = 0
        self.docNumPagesMsg = "Not processed"
        self.docIsEncrypted = False
        self.docIsEncryptedMsg = "Not processed"

        # Attempt to process the file
        # If the long file path flag is triggered, set error messages and do not process, unless the ignore max path flag is also on.
        if self.docPathFullLenFlag and not self.docIgnoreMaxPathLenFlag:
            self.docSizeMsg = "Cannot read size - long file path red flag triggered"
            self.docNumPagesMsg = "Cannot count number of pages - long file path red flag triggered"
            self.docIsEncryptedMsg = "Cannot check if encrypted - long file path red flag triggered"         

        # Otherwise, if the long file path flag hasn't been triggered (or if it has but the ignore max path flag has is on), continue to process the file.
        else:
            # Get the Document file size.
            self.checkDocSize()

            # Check the Document file size against the max processing size threshold. The max processing size is taken from the constructor (optional). If this is set to 0, no limit is applied. The default is 0. If the Document file size is greater than the threshold, set the red flag for exceeding this.      
            if self.docMaxSizeForProcessing != 0 and self.docSize > self.docMaxSizeForProcessing:
                self.docSizeMsg = "File exceeds maximum size threshold for processing"
                self.docNumPagesMsg = "Cannot count number of pages - file exceeds maximum size threshold for processing"
                self.docIsEncryptedMsg = "Cannot check if encrypted - file exceeds maximum size threshold for processing"
            else:   
                self.checkDocIsEncrypted()
                self.checkDocNumPages()


        """
        character counter for txt and doc files

        TO DO - consider if helpful?

        create a character counter routine to read number of characters in text based files such as TXT and Word docs

        """

        # Initialise the meta data labels (used for outputs)
        self.setDocDataLabels()          

    """
    END OF init METHOD
    """

    """
    METHOD
    Set this Document.
    Re-initialises this Document - used to change the target file when the Document object is retained
    """
    def setDoc(self, rootDir, fullPath, indexNum=0, indexDigitsSize = 5, maxSizeForProcessing = 0):
        self.__init__(rootDir, fullPath, indexNum=0, indexDigitsSize = 5, maxSizeForProcessing = 0)

    """
    METHOD 
    Setup Document path based metadata.
    Sets name and path details by getting them from the document path.
    """
    def setDocPathDetails(self):

        # Check if the path contains a comma; if it does, flag this
        if self.docPathFull.find(",") > -1:
            self.docPathCommaFlag = True
        else:
            self.docPathCommaFlag = False
        
        # Set the doc path without the root - this is done by replacing a substring of the root with nothing
        self.docPathNoRoot = self.docPathFull.replace(self.docRootDirPath+"\\","",-1)

        # Set the full path length and check if it exceeds the maximum.
        # The standard Windows full path limit (path and file name) is 260 characters.
        # Sets a red flag if the full path is greater than this.
        self.docPathFullLen = len(self.docPathFull)
        if self.docPathFullLen > 260:
            self.docPathFullLenFlag = True
        else:
            self.docPathFullLenFlag = False     
        
        # Try and find the path only.
        # Expression includes end slash to avoid throwing off character counters by 1.
        # Assumes path separators are either single forward or back slashes. 
        # This regex pattern assumes there is a file extension.
        # If the path only isn't found, try a regex search pattern that assumes there isn't a file extension.
        try:
            self.docPathOnly = re.search(
                r"(^.*\\|/)(.*)(\.)(.*$)",
                self.docPathFull).group(1)
        except:
            self.docPathOnly = re.search(
                r"(^.*\\|/)(.*$)",
                self.docPathFull).group(1)

        # Set the length of the path only
        # The standard Windows path only limit (path without file name) is 248 characters.
        # Sets a red flag if the full path is greater than this.
        self.docPathOnlyLen = len(self.docPathOnly)        
        if self.docPathOnlyLen > 248:
            self.docPathOnlyLenFlag = True
        else:
            self.docPathOnlyLenFlag = False

        # Regex for the full document name (i.e. with extension)
        # Assumes path separators are either single forward or back slashes. 
        self.docNameFull = re.search(
            r"(^.*)(\\|/)(.*$)",
            self.docPathFull).group(3)

        # Regex for the document name without extension. This regex pattern assumes there is a file extension.
        # Assumes path separators are either single forward or back slashes. 
        # If the file name isn't found, try a regex search pattern that assumes there isn't a file extension instead.
        try:
            self.docNameNoExt = re.search(
                r"(^.*)(\\|/)(.*)(\.)(.*$)",
                self.docPathFull).group(3)
        except:
            self.docNameNoExt = re.search(
                r"(^.*)(\\|/)(.*$)",
                self.docPathFull).group(3)

        # Setup variations of the full file name with and without file extension where
        # - commas are replaced with underscores
        self.docNameFullClean = ""
        self.docNameNoExtClean = ""
        self.docNameFullClean = re.sub(r"[,]","_",self.docNameFull)
        #self.docNameFullClean = re.sub(r"[&]","and",self.docNameFullClean)
        self.docNameNoExtClean = re.sub(r"[,]","_",self.docNameNoExt)
        #self.docNameNoExtClean = re.sub(r"[&]","and",self.docNameNoExtClean)

        # Regex for the document extension. If none is found, set extension as unknown.
        try:
            self.docExtension = re.search(r"(^.*)(\.)(.*$)",self.docNameFull).group(3).lower()
        except:
            self.docExtension = "unknown"
              
    """
    METHOD
    Get the size of this document in bytes.
    First checks that the long file path flag hasn't been triggered. If it has, handles this as an error and sets the size to 0.
    """
    def checkDocSize(self):
        try:
            self.docSize = os.path.getsize(self.docPathFull)
            self.docSizeMsg = "Successfully read file size"
        except Exception as error:
            self.docSizeMsg = error
        
    """
    METHOD
    Checks if this Document is encrypted.
    Encryption can only reliably be detected on PDFs. If the file is not a PDF, the file encryption status is flagged as unknown.

    Technical Notes:
    The pywin32 library uses VBA commands to interact with Word files. Unfortunately, the commands to check if a file is password encyrypted need to open the file and, where there is a password, the window will hang on the password prompt. This pauses the script and, therefore, pywin32 cannot be used. Available commands cannot close the prompt nor avoid opening the file altogether. As such, we only check for encryption on PDFs the PyPDF2 lib.

    """
    def checkDocIsEncrypted(self):
        isEncryptedMessage = "Warning - file is encrypted"
        isNotEncryptedMessage = "File is not encrypted"
        isUnknownEncryptedMessage = "Unknown - file cannot be checked for encryption"
        try:
            # If the Document is a PDF
            if self.docExtension == "pdf":            
                    pdfObj = pdfReader.PdfFileReader(open(self.docPathFull, 'rb'))
                    self.docIsEncrypted = pdfObj.isEncrypted
                    if  self.docIsEncrypted:
                        self.docIsEncryptedMsg = isEncryptedMessage
                    else:
                        self.docIsEncryptedMsg = isNotEncryptedMessage                            

            # If the Document is not a recognised file format for encryption checking
            else:
                self.docIsEncrypted = False
                self.docIsEncryptedMsg = isUnknownEncryptedMessage

        except Exception as error:
            self.docIsEncrypted = False
            self.docIsEncryptedMsg = error   
       
    """
    METHOD
    Check the number of pages of this Document.
    Parse this Document to get the number of pages if it is a file format that has pages as a object or concept. Supported file types:
        - PDFs
        - tiffs
    A default error message is set indicating 'success' in processing to begin with.
    """
    def checkDocNumPages(self):
        numPagesSuccessMessage = "Successfully read number of pages"

        # If the file is a PDF, process it using PyPdf2 library
        # PyPdf2 flattens PDFs in order to count the number of pages. If the file is encrypted, it will instead rely on the page count metadata of the file itself rather than an actual count. It is implied by comments in the PyPDF2 lib that the latter is less reliable.
        if self.docExtension == "pdf":
            try:
                pdfObj = pdfReader.PdfFileReader(open(self.docPathFull, 'rb'))
                if  self.docIsEncrypted:
                    self.docNumPagesMsg = "Warning - document is an encrypted PDF - file page count metadata relied on instead of flattened page count"
                    self.docNumPages = pdfObj.getNumPages()
                else:
                    self.docNumPagesMsg = numPagesSuccessMessage
                    self.docNumPages = pdfObj.getNumPages()
            except Exception as error:
                self.docNumPagesMsg = error

        # If the file is a TIFF image, count the number of image frames
        elif self.docExtension in ["tif","tiff"]:
            
            try:
                tiffObj = Image.open(self.docPathFull)
                #tiffObj.load(0
                self.docNumPages = tiffObj.n_frames
            except Exception as error:
                self.docNumPagesMsg = error

        else:
            self.docNumPagesMsg = "Cannot process - not a supported file type for page counting"

            """
            End of num pages processor
            """

    """
    METHOD
    Set up a list of default text labels used for output
    """
    def setDocDataLabels(self):
        self.docIndexNumLabel = 'Index (Number)'
        self.docIndexDigitsLabel = 'Index (Digits)'
        
        self.docNameFullLabel = 'File Name (With Extension)'
        self.docNameNoExtLabel = 'File Name (No Extension)'
        self.docExtensionLabel = 'File Extension'
        self.docNameFullCleanLabel = 'Clean File Name (With Extension)'
        self.docNameNoExtCleanLabel = 'Clean File Name (No Extension)'
        
        self.docPathFullLabel = 'Full Path'
        self.docSizeLabel = 'File Size (Bytes)'
        self.docSizeMsgLabel = 'File Size Check Log'
        
        self.docPathCommaFlagLabel = 'File Path Contains Comma'
        self.docPathFullLenFlagLabel = 'Full Path Length Red Flag'
        self.docPathOnlyLenFlagLabel = 'Path Only Length Red Flag'
        self.docMaxSizeForProcessingSkipFlagLabel = 'Large File Size Flag'
        
        self.docPathFullLenLabel = 'Full Path Length'
        self.docPathOnlyLenLabel = 'Path Only Length'
        
        self.docNumPagesLabel = '(PDFs Only) Number of Pages' 
        self.docIsEncryptedLabel = '(PDFs Only) Is the File Encrypted?'
        self.docIsEncryptedMsgLabel = '(PDFs Only) File Encryption Check Log'
        self.docNumPagesMsgLabel = '(PDFs Only) Number of Pages Log'
        
        self.docPathOnlyLabel = 'Path Only'
        self.docPathNoRootLabel = 'Path Without Root'


    """
    METHOD
    Returns a dictionary object representation of this Document with pre-set meaningful key labels.
    """
    def getDocVarDict(self):      
        docVarDict = {
            self.docIndexNumLabel : self.docIndexNum,
            self.docIndexDigitsLabel : self.docIndexDigits,
            
            self.docNameFullLabel : self.docNameFull,
            self.docNameNoExtLabel : self.docNameNoExt,  
            self.docExtensionLabel : self.docExtension,
            self.docNameFullCleanLabel : self.docNameFullClean,
            self.docNameNoExtCleanLabel : self.docNameNoExtClean,
            
            self.docPathFullLabel: self.docPathFull,
            self.docSizeLabel : self.docSize,
            self.docSizeMsgLabel : self.docSizeMsg,
            
            self.docPathCommaFlagLabel: self.docPathCommaFlag,
            self.docPathFullLenFlagLabel : self.docPathFullLenFlag,
            self.docPathOnlyLenFlagLabel : self.docPathOnlyLenFlag,
            self.docMaxSizeForProcessingSkipFlagLabel : self.docMaxSizeForProcessingSkipFlag,
            
            self.docPathFullLenLabel : self.docPathFullLen,
            self.docPathOnlyLenLabel: self.docPathOnlyLen,

            self.docNumPagesLabel: self.docNumPages,
            self.docIsEncryptedLabel: self.docIsEncrypted,
            self.docIsEncryptedMsgLabel : self.docIsEncryptedMsg,
            self.docNumPagesMsgLabel : self.docNumPagesMsg,
            
            self.docPathOnlyLabel: self.docPathOnly,
            self.docPathNoRootLabel: self.docPathNoRoot,
        }
        return docVarDict

    """
    METHOD
    Returns a list representation of this Document.
    """
    def getDocVarList(self):      
        docVarList = [
            self.docIndexNum,
            self.docIndexDigits,
            
            self.docNameFull,
            self.docNameNoExt,            
            self.docExtension,
            self.docNameFullClean,
            self.docNameNoExtClean,

            self.docPathFullLenFlag,
            self.docPathOnlyLenFlag, 
            self.docMaxSizeForProcessingSkipFlag,          
            
            self.docPathCommaFlag,
            self.docPathFull,            
            self.docSize,
            self.docSizeMsg,

            self.docPathFullLen,
            self.docPathOnlyLen,

            self.docNumPages,
            self.docIsEncrypted,
            self.docIsEncryptedMsg,            

            self.docNumPagesMsg,
            self.docPathOnly,
            self.docPathNoRoot,
        ]
        return docVarList

    """
    METHOD
    Prints out the vars of this Document using pre-set meaningful labels.
    """
    def printDocVars(self):
        print(self.docIndexNumLabel + ": " + str(self.docIndexNum))
        print(self.docIndexDigitsLabel + ": " + self.docIndexDigits)

        print(self.docNameFullLabel + ": " + self.docNameFull)
        print(self.docNameNoExtLabel + ": " + self.docNameNoExt)            
        print(self.docExtensionLabel + ": " + self.docExtension)
        print(self.docNameFullCleanLabel + ": " + self.docNameFullClean)
        print(self.docNameNoExtCleanLabel + ": " + self.docNameNoExtClean)      
        
        print(self.docPathFullLenFlagLabel + ": " + str(self.docPathFullLenFlag))
        print(self.docPathOnlyLenFlagLabel + ": " + str(self.docPathOnlyLenFlag))
        print(self.docMaxSizeForProcessingSkipFlagLabel + ": " + str(self.docMaxSizeForProcessingSkipFlag))

        print(self.docPathCommaFlagLabel + ": " + self.docPathCommaFlag)
        print(self.docPathFullLabel + ": " + self.docPathFull)
        print(self.docSizeLabel + ": " + str(self.docSize))
        print(self.docSizeMsgLabel + ": " + self.docSizeMsg)
        
        print(self.docPathFullLenLabel + ": " + str(self.docPathFullLen))
        print(self.docPathOnlyLenLabel + ": " + str(self.docPathOnlyLen))
        print(self.docNumPagesLabel + ": " + str(self.docNumPages))
        print(self.docIsEncryptedLabel + ": " + str(self.docIsEncrypted))
        print(self.docIsEncryptedMsgLabel + ": " + self.docIsEncryptedMsg)
        print(self.docNumPagesMsgLabel + ": " + self.docNumPagesMsg)
        print(self.docPathOnlyLabel + ": " + self.docPathOnly)
        print(self.docPathNoRootLabel + ": " + self.docPathNoRoot)