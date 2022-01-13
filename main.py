#Authored by Seaghan Ennis
#Last Update: Jan 13, 2022

import docx

def main():
    #enter script
    print("Starting formatter")

    #opens the document and snapshots it
    inputFile = "example.docx"
    print("Opening: ", inputFile)
    inDoc = docx.Document(inputFile)

    #setup outDoc
    outDoc = docx.Document()
    
    for i in inDoc.paragraphs:
        print(i.text)
    


if __name__ == '__main__':
    #start script
    main()