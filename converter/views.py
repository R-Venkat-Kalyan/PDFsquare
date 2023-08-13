import pythoncom
pythoncom.CoInitialize()

from pdf2image import convert_from_path
import subprocess
from django.shortcuts import render
from django.http import HttpResponse
from docx import Document
import os,io,img2pdf,magic,fitz,tempfile,zipfile,PyPDF2,ppt2pdf
from PyPDF2 import PdfReader,PdfWriter,PdfMerger
from pdf2docx import Converter
from docx2pdf import convert
from PIL import Image
from pdf2image import convert_from_path
from pptx import Presentation
from pathlib import Path
import pythoncom
import win32com.client
from fpdf import FPDF
from django.core.files.storage import FileSystemStorage
from tempfile import TemporaryDirectory
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from io import BytesIO
import aspose.slides as slides
from tempfile import NamedTemporaryFile

# website landing page / home page
def home(request):
    return render(request,'home.html')


# ABOUT US
def AboutUs(request):
    return render(request,'AboutUs.html')


#Contact Us Form
def ContactUs(request):
    return render(request, 'ContactUS.html')


# Feedback Form
def UserFeedback(request):
    return render(request,'Feedback.html')


#pdf to word converter file upload and download page(card1)
def PdfToWordConverter(request):
    if request.method == 'POST' and 'file' in request.FILES:
        selected_file = request.FILES['file']

        # Save the selected file
        file_path = os.path.join("media", selected_file.name)

        fs = FileSystemStorage()
        fs.save(file_path, selected_file)

        # Convert the PDF to Word
        pdf_path = os.path.join(fs.location, file_path)
        docx_path = os.path.join(fs.location, file_path.replace(".pdf", ".docx"))

        # Perform the conversion
        try:
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()

            # Serve the converted DOCX file for download
            with open(docx_path, 'rb') as f:
                response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                response['Content-Disposition'] = 'attachment; filename=' + os.path.basename(docx_path)
                return response

        finally:
            # Delete the temporary files
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
            if os.path.exists(docx_path):
                os.remove(docx_path)

    return render(request, 'pdf_to_word_converter.html')


#Word to Pdf converter file upload and download page(card2)
def DocxToPdfConverter(request):
    if request.method == 'POST' and 'file' in request.FILES:
        selected_file = request.FILES['file']

        # Save the selected file
        file_path = os.path.join("media", selected_file.name)

        fs = FileSystemStorage()
        fs.save(file_path, selected_file)

        # Initialize COM environment
        pythoncom.CoInitialize()

        # Convert the DOCX to PDF
        docx_path = os.path.join(fs.location, file_path)
        pdf_path = os.path.join(fs.location, file_path.replace(".docx", ".pdf"))

        # Perform the conversion
        try:
            convert(docx_path, pdf_path)

            # Serve the converted PDF file for download
            with open(pdf_path, 'rb') as f:
                response = HttpResponse(f.read(), content_type='application/pdf')
                response['Content-Disposition'] = 'attachment; filename=' + os.path.basename(pdf_path)
                return response

        finally:
            # Delete the temporary files
            if os.path.exists(docx_path):
                os.remove(docx_path)
            if os.path.exists(pdf_path):
                os.remove(pdf_path)

    return render(request, 'word_to_pdf_converter.html')


# Image to PDF converter file upload and download page(card3)
def ImgToPdfConverter(request):
    if request.method == 'POST' and 'file' in request.FILES:
        selected_file = request.FILES['file']

        # Check if the uploaded file is an image
        try:
            image = Image.open(selected_file)
        except (OSError, Image.UnidentifiedImageError):
            return HttpResponse('Invalid image file')

        # Create a BytesIO object and save the file content
        file_stream = io.BytesIO()
        for chunk in selected_file.chunks():
            file_stream.write(chunk)
        file_stream.seek(0)

        # Convert the image to PDF
        output_filename = selected_file.name.split('.')[0] + '.pdf'
        with open(output_filename, 'wb') as f:
            f.write(img2pdf.convert(file_stream))

        # Serve the converted PDF file for download
        with open(output_filename, 'rb') as f:
            response = HttpResponse(f.read(), content_type='application/pdf')
            response['Content-Disposition'] = 'attachment; filename=' + output_filename

        # Remove the converted PDF file from the server
        os.remove(output_filename)

        return response

    return render(request, 'image_to_pdf_converter.html')


# PDF to IMAGE converter file upload and download page(card4)
def PdfToImgConverter(request):
    if request.method == 'POST' and 'file' in request.FILES:
        selected_file = request.FILES['file']

        # Check if the uploaded file is a PDF
        if not selected_file.name.endswith('.pdf'):
            return HttpResponse('Invalid PDF file')

        # Read the selected file content into a BytesIO object
        pdf_data = io.BytesIO()
        for chunk in selected_file.chunks():
            pdf_data.write(chunk)
        pdf_data.seek(0)

        # Convert the PDF to images
        images = convert_to_images(pdf_data)

        # Create a BytesIO object for the zip file
        zip_memory = io.BytesIO()

        # Create the zip file and add images to it
        with zipfile.ZipFile(zip_memory, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for i, image in enumerate(images):
                image_data = io.BytesIO()
                image.save(image_data, format='JPEG')
                image_filename = f'{i}.jpg'

                # Add the image data to the zip
                zipf.writestr(image_filename, image_data.getvalue())

        # Serve the zip file for download
        zip_memory.seek(0)
        response = HttpResponse(zip_memory.read(), content_type='application/zip')
        response['Content-Disposition'] = 'attachment; filename=converted_images.zip'
        return response

    return render(request, 'pdf_to_image_converter.html')


def convert_to_images(pdf_data):
    images = []
    doc = fitz.open(stream=pdf_data, filetype='pdf')
    for page_number in range(doc.page_count):
        page = doc.load_page(page_number)
        pix = page.get_pixmap()
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(image)
    doc.close()
    return images

# pdf compresser(card5)
def PdfCompresser(request):
    if request.method == 'POST' and 'file' in request.FILES:
        selected_file = request.FILES['file']

        # Save the selected file temporarily on the server
        temp_filename = f'temp_{selected_file.name}'
        with open(temp_filename, 'wb') as f:
            f.write(selected_file.read())

        # Create a temporary file for the compressed PDF
        _, output_filename = tempfile.mkstemp(suffix='.pdf')

        # Perform PDF compression
        with open(temp_filename, 'rb') as input_file:
            pdf_reader = PdfReader(input_file)

            # Check if the PDF is encrypted
            if pdf_reader.is_encrypted:
                # If encrypted, try to decrypt using an empty password
                pdf_reader.decrypt('')

            pdf_writer = PdfWriter()

            for page_number in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_number]
                page.compress_content_streams()  # Compress content streams
                pdf_writer.add_page(page)

            with open(output_filename, 'wb') as output_file:
                pdf_writer.write(output_file)

        # Remove the temporary files from the server
        os.remove(temp_filename)

        # Serve the compressed PDF file for download
        with open(output_filename, 'rb') as f:
            response = HttpResponse(f.read(), content_type='application/pdf')
            response['Content-Disposition'] = 'attachment; filename=compressed.pdf'
            return response

    return render(request, 'pdf_compresser.html')

#pdf merger(card6)
def MergePdfs(request):
    if request.method == 'POST':
        pdf_files = request.FILES.getlist('pdf_files[]')

        # Create a PDF merger object
        merger = PdfMerger()

        for pdf_file in pdf_files:
            # Merge each PDF file
            merger.append(pdf_file)

        # Create a response object
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="merged.pdf"'

        # Write the merged PDF content into the response
        merger.write(response)

        # Close the merger
        merger.close()

        # Return the response
        return response

    return render(request, 'pdf_merger.html')


#ppt to pdf converter(card7)
def PptToPdfConverter(request):
    if request.method == 'POST' and 'file' in request.FILES:
        ppt_file = request.FILES['file']

        # Load presentation
        ppt_bytes = ppt_file.read()
        pres = slides.Presentation(io.BytesIO(ppt_bytes))

        # Create PDF options
        options = slides.export.PdfOptions()

        # Set desired compliance and save as PDF
        options.compliance = slides.export.PdfCompliance.PDF_A1A
        pdf_bytes = io.BytesIO()
        pres.save(pdf_bytes, slides.export.SaveFormat.PDF, options)
        pdf_bytes.seek(0)

        # Prepare the file for download
        response = HttpResponse(pdf_bytes, content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="converted_pdf.pdf"'
        return response

    return render(request, 'ppt_to_pdf_converter.html')

def secure_pdf(input_file, password):
    pdf_reader = PyPDF2.PdfReader(input_file)
    pdf_writer = PyPDF2.PdfWriter()

    for page_num in range(len(pdf_reader.pages)):
        pdf_writer.add_page(pdf_reader.pages[page_num])

    pdf_writer.encrypt(password)

    output_file = io.BytesIO()
    pdf_writer.write(output_file)
    output_file.seek(0)

    return output_file

# pdf encrypter(card8)
def PdfEncrypter(request):
    if request.method == 'POST' and 'file' in request.FILES:
        selected_file = request.FILES['file']
        password = request.POST['password']

        # Secure the PDF file with the user-provided password
        encrypted_file = secure_pdf(selected_file, password)

        # Prepare the encrypted PDF file for download
        response = HttpResponse(encrypted_file.getvalue(), content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="secured_file.pdf"'

        # Close the BytesIO object
        encrypted_file.close()

        return response

    return render(request, 'pdf_encrypter.html')