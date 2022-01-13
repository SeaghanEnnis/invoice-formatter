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
    style = outDoc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = docx.shared.Pt(12)
    paragraph_format = outDoc.styles['Normal'].paragraph_format
    paragraph_format.space_before = docx.shared.Pt(0)
    paragraph_format.space_after = docx.shared.Pt(0)
    paragraph_format.line_spacing = 1

    #Read inDoc to produce outDoc
    for p in inDoc.paragraphs:

        # Remove space
        for char in p.text:
            if char == " ":
                p.text = p.text[1:]
            else:
                break

        outDoc.add_paragraph(p.text)

    outDoc.save("exampleOutput.docx")

    print("Finished and saved")

if __name__ == '__main__':
    #start script
    main()