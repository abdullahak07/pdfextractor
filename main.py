from appJar import gui
import PyPDF2
from PyPDF2 import PdfFileWriter, PdfFileReader
from pathlib import Path
from PyPDF2 import PdfFileMerger

import pytesseract
from PIL import Image
from wand.image import Image as wi
import os
# Define all the functions needed to process the files


def split_pages(input_file, page_range, out_file):
    """ Take a pdf file and copy a range of pages into a new pdf file

    Args:
        input_file: The source PDF file
        page_range: A string containing a range of pages to copy: 1-3,4
        out_file: File name for the destination PDF
    """
    output = PdfFileWriter()
    input_pdf = PdfFileReader(open(input_file, "rb"))
    output_file = open(out_file, "wb")

    # https://stackoverflow.com/questions/5704931/parse-string-of-integer-sets-with-intervals-to-list
    page_ranges = (x.split("-") for x in page_range.split(","))
    range_list = [i for r in page_ranges for i in range(int(r[0]), int(r[-1]) + 1)]

    for p in range_list:
        # Need to subtract 1 because pages are 0 indexed
        try:
            output.addPage(input_pdf.getPage(p - 1))
        except IndexError:
            # Alert the user and stop adding pages
            app.infoBox("Info", "Range exceeded number of pages in input.\nFile will still be saved.")
            break
    output.write(output_file)

    if(app.questionBox("File Save", "Output PDF saved. Do you want to quit?")):
        app.stop()


def validate_inputs(input_file, output_dir, range, file_name):
    """ Verify that the input values provided by the user are valid

    Args:
        input_file: The source PDF file
        output_dir: Directory to store the completed file
        range: File A string containing a range of pages to copy: 1-3,4
        file_name: Output name for the resulting PDF

    Returns:
        True if error and False otherwise
        List of error messages
    """
    errors = False
    error_msgs = []

    # Make sure a PDF is selected
    if Path(input_file).suffix.upper() != ".PDF":
        errors = True
        error_msgs.append("Please select a PDF input file")

    # Make sure a range is selected
    if len(range) < 1:
        errors = True
        error_msgs.append("Please enter a valid page range")

    # Check for a valid directory
    if not(Path(output_dir)).exists():
        errors = True
        error_msgs.append("Please Select a valid output directory")

    # Check for a file name
    if len(file_name) < 1:
        errors = True
        error_msgs.append("Please enter a file name")

    return(errors, error_msgs)


def press(button):
    """ Process a button press

    Args:
        button: The name of the button. Either Process of Quit
    """
    if button == "Process":
        src_file = app.getEntry("Input_File")
        dest_dir = app.getEntry("Output_Directory")
        page_range = app.getEntry("Page_Ranges")
        out_file = app.getEntry("Output_name")
        errors, error_msg = validate_inputs(src_file, dest_dir, page_range, out_file)
        if errors:
            app.errorBox("Error", "\n".join(error_msg), parent=None)
        else:
            split_pages(src_file, page_range, Path(dest_dir, out_file))
    elif button == "PDF Extractor":
        print('-----Starting Convertion-----')
        src_file = app.getEntry("Input_File")
        out_file = app.getEntry("Output_name")

         # creating a pdf file object
        pdfFileObj = open(src_file, 'rb')

            # creating a pdf reader object
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

            # printing number of pages in pdf file
        print(pdfReader.numPages)

            # creating a page object
        pageObj = pdfReader.getPage(0)
        f = open(out_file, "w+")
        f.write(pageObj.extractText())
            # extracting text from page
        print(pageObj.extractText())
        f.write(text)


            # closing the pdf file object
        pdfFileObj.close()
        f.close()
        print('----- Convertion Done-----')
        if (app.questionBox("File Save", "Output PDF saved. Do you want to quit?")):
            app.stop()
    elif button == "PDF Merger":
        print('----- Merging Done -----')


        src_file = app.getEntry("Input_File")
        merge_file= app.getEntry("Merge_File")
        out_file = app.getEntry("Output_name")

        pdfs = [src_file, merge_file]

        merger = PdfFileMerger()

        for pdf in pdfs:
            merger.append(pdf)

        merger.write(out_file)
        merger.close()
        if (app.questionBox("File Save", "Output PDF saved. Do you want to quit?")):
            app.stop()
            print('----- Merging Done -----')
    elif button == "ImagePdfExtractor":
        src_file = app.getEntry("Input_File")
        out_file = app.getEntry("Output_name")
        print('-----Starting Convertion-----')

        pdf = wi(filename=src_file, resolution=300)
        pdfimage = pdf.convert("jpeg")
        i = 1
        for img in pdfimage.sequence:
            page = wi(image=img)
            v=page.save(filename=str(i) + ".jpg")
            i += 1

            pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"

        for x in range(v):


            img = Image.open("C:\\OC\\{0}.jpg".format(x))

            text = pytesseract.image_to_string(img)
            print(text)
            f = open(out_file,'a+')
            f.write(text)
        print('-----Done-----')
        if (app.questionBox("File Save", "Output PDF saved. Do you want to quit?")):
            app.stop()
    elif button == "PDF Splitter":
        src_file = app.getEntry("Input_File")
        out_file = app.getEntry("Output_name")

        def PDFsplit(pdf, splits):
            # creating input pdf file object
            pdfFileObj = open(pdf, 'rb')

            # creating pdf reader object
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

            # starting index of first slice
            start = 0

            # starting index of last slice
            end = splits[0]

            for i in range(len(splits) + 1):
                # creating pdf writer object for (i+1)th split
                pdfWriter = PyPDF2.PdfFileWriter()

                # output pdf file name
                outputpdf = pdf.split('.pdf')[0] + str(i) + '.pdf'

                # adding pages to pdf writer object
                for page in range(start, end):
                    pdfWriter.addPage(pdfReader.getPage(page))

                    # writing split pdf pages to pdf file
                with open(outputpdf, "wb") as f:
                    pdfWriter.write(f)

                    # interchanging page split start position for next split
                start = end
                try:
                    # setting split end position for next split
                    end = splits[i + 1]
                except IndexError:
                    # setting split end position for last split
                    end = pdfReader.numPages

                    # closing the input pdf file object
            pdfFileObj.close()
            if (app.questionBox("File Save", "Output PDF saved. Do you want to quit?")):
                app.stop()

        def main():
            # pdf file to split
            pdf = src_file
            # split page positions
            splits = [2, 4]

            # calling PDFsplit function to split pdf
            PDFsplit(pdf, splits)

        if __name__ == "__main__":
            # calling the main function
            main()
    elif button == "PDF Rotater":
        src_file = app.getEntry("Input_File")


        def PDFrotate(origFileName, newFileName, rotation):

            # creating a pdf File object of original pdf
            pdfFileObj = open(origFileName, 'rb')

            # creating a pdf Reader object
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

            # creating a pdf writer object for new pdf
            pdfWriter = PyPDF2.PdfFileWriter()

            # rotating each page
            for page in range(pdfReader.numPages):
                # creating rotated page object
                pageObj = pdfReader.getPage(page)
                pageObj.rotateClockwise(rotation)

                # adding rotated page object to pdf writer
                pdfWriter.addPage(pageObj)

                # new pdf file object
            newFile = open(newFileName, 'wb')

            # writing rotated pages to new file
            pdfWriter.write(newFile)

            # closing the original pdf file object
            pdfFileObj.close()

            # closing the new pdf file object
            newFile.close()

        def main():

            # original pdf file name
            origFileName = src_file
            out_file = app.getEntry("Output_name")
            # new pdf file name
            newFileName = out_file

            # rotation angle
            rotation = 270

            # calling the PDFrotate function
            PDFrotate(origFileName, newFileName, rotation)

        if __name__ == "__main__":
            # calling the main function
            main()
    else:
        app.stop()

# Create the GUI Window
app = gui("PDF Splitter", useTtk=True)
app.setTtkTheme("default")
app.setSize(500, 200)

# Add the interactive components
app.addLabel("Choose Source PDF File")
app.addFileEntry("Input_File")

app.addLabel("Choose Merge PDF File")
app.addFileEntry("Merge_File")
app.addLabel("Select Output Directory")
app.addDirectoryEntry("Output_Directory")

app.addLabel("Output file name")
app.addEntry("Output_name")

app.addLabel("Page Ranges: 1,3,4-10")
app.addEntry("Page_Ranges")

# link the buttons to the function called press
app.addButtons(["PDF Extractor","PDF Merger","PDF Rotater","ImagePdfExtractor", "Quit"], press)



# start the GUI
app.go()
