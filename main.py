#Authored by Seaghan Ennis
#Last Update: Jan 13, 2022

import docx

def main():
    #enter script
    print("Starting formatter")

    #opens the document and snapshots it
    inputFile = "example.docx"
    print("Opening: ", inputFile)
    doc = docx.Document(inputFile)
    
    for i in doc.paragraphs:
        print(i.text)


if __name__ == '__main__':
    #start script
    main()