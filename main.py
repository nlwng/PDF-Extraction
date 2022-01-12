import io
import os

import camelot as cam
import fitz
import pikepdf
import pytesseract
from pdf2image import convert_from_path
from PIL import Image


class PDFEXTRACTION:
    """
    A Class is created to extract the following things from a pdf
    Text,Image,Table

    """

    def __init__(self, filename, pass_word):
        """
        A method is created for initialize the filename and the password

        """
        self.filename = filename
        self.pass_word = pass_word

    def decrypt(self):
        """
        A method is created for decrypting a pdf

        """
        password = self.pass_word
        with pikepdf.open(self.filename, password=password) as pdf:
            pdf.save('output.pdf')

    def extracting_images(self):
        """
        A method is created for extracting all the images present in the pdf

        """

        image_dir = './image'
        # open the file
        pdf_file = fitz.open("output.pdf")
        len_pdf = len(pdf_file)
        # iterate over PDF pages
        for page_index in range(len_pdf):
            # get the page itself

            page = pdf_file[page_index]
            image_list = page.getImageList()
            # printing number of images found in this page
            if image_list:
                print(
                    f"[+] Found {len(image_list)} images in page {page_index}"
                    )
            else:
                print("[!] No images found on page", page_index)
            for image_index, img in enumerate(page.getImageList(), start=1):
                # get the XREF of the image
                xref = img[0]
                # extract the image bytes
                base_image = pdf_file.extract_image(xref)
                image_bytes = base_image["image"]
                # get the image extension
                image_ext = base_image["ext"]
                # load it to PIL
                image = Image.open(io.BytesIO(image_bytes))

                # save it to local disk
                file_name = f"image{page_index+1}_{image_index}.{image_ext}"
                filePath = os.path.join(image_dir, file_name)

                image.save(open(filePath, "wb"))

    def extracting_text(self):

        """
        The method is extracting all the text from the pdf
        And saved in a text file

        """
        path_teseract = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
        pytesseract.pytesseract.tesseract_cmd = path_teseract
        # Opening pages from path of pdf
        pages = convert_from_path(
            "output.pdf", 500,
            poppler_path=r"C:\\Users\\hp\\Downloads\\poppler-0.68.0\\bin"
            )

        image_counter = 1

        for page in pages:
            # Saving pages as image
            filename = "page_"+str(image_counter)+".jpg"
            page.save(filename, 'JPEG')

            image_counter = image_counter + 1

        filelimit = image_counter-1

        outfile = "out_text.txt"

        f = open(outfile, "a")

        for i in range(1, filelimit + 1):

            filename = "page_"+str(i)+".jpg"

            # Taking Text from images
            text = str(((pytesseract.image_to_string(Image.open(filename)))))

            text = text.replace('-\n', '')

            f.write(text)

        f.close()

    def table_pdf(self):
        """
        The method is extracting all the tables present in the PDF
        And saving in a CSV file

        """

        table = cam.read_pdf("output.pdf", pages='1', flavor='stream')
        table.export('next.csv', f='csv', compress=False)
        # Extracting all the tables


FILENAME = "eStatement9511IN_2021-11-23_17_23.pdf"

# Give Password if the pdf is encrypted, Otherwise leave it.
PASSWORD = "LOKE180874"

obj = PDFEXTRACTION(FILENAME, PASSWORD)

obj.decrypt()
obj.extracting_images()
obj.extracting_text()
obj.table_pdf()
os.remove("output.pdf")
