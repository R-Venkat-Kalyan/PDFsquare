
from django.shortcuts import render,redirect
from django.http import HttpResponse
from docx import Document
import os,io,img2pdf,fitz,tempfile,zipfile,PyPDF2
from pdf2docx import Converter
from docx2pdf import convert
from PIL import Image
import PyPDF2
from django.core.files.storage import FileSystemStorage
import aspose.slides as slides
from .models import Register
from django.db.models import Q
from django.contrib import messages

# website landing page / home page
def home(request):
    return render(request,'BCT_Home.html')

def login(request):
    return render(request,'Login_form.html')

def register(request):
    if request.method == 'POST':
        if request.POST.get('fn') and request.POST.get('em')and request.POST.get('pwd')and request.POST.get('re'):
            post = Register()
            post.Name = request.POST.get('fn')
            post.E_mail = request.POST.get('em')
            post.password = request.POST.get('pwd')
            post.Re_password = request.POST.get('re')
            post.save()
    return render(request,'Registration_form.html')

def introduction(request):

    email = request.POST['em']
    pwd = request.POST['pwd']

    flag = Register.objects.filter(Q(E_mail=email) & Q(password=pwd))

    if flag:
        messages.success(request, 'Login Successful')
        return render(request, 'BCT_HOME1.html')
    else:
        messages.error(request, 'Invalid Login credentials...!')
        return redirect('login')

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

        # Convert the DOCX to PDF
        docx_path = os.path.join(fs.location, file_path)
        pdf_path = os.path.join(fs.location, file_path.replace(".docx", ".pdf"))

        try:
            # Use docx2pdf to convert DOCX to PDF
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
        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
            temp_file.write(selected_file.read())
            temp_filename = temp_file.name

        # Create a temporary file for the compressed PDF
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as output_file:
            output_filename = output_file.name

        # Perform PDF compression
        pdf_document = fitz.open(temp_filename)
        pdf_document.save(output_filename, deflate=True, garbage=3, clean=True)
        pdf_document.close()  # Close the PDF document

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

        # Create a temporary file to store the merged PDF
        merged_pdf_path = os.path.join(tempfile.gettempdir(), 'merged.pdf')
        pdf_document = fitz.open()

        try:
            for pdf_file in pdf_files:
                # Save the selected PDF file temporarily
                with tempfile.NamedTemporaryFile(delete=False) as temp_file:
                    temp_file.write(pdf_file.read())
                    temp_filename = temp_file.name

                # Add the PDF to the merged document
                pdf_document.insert_pdf(fitz.open(temp_filename))

                # Remove the temporary PDF file
                os.remove(temp_filename)

            # Save the merged PDF
            pdf_document.save(merged_pdf_path, garbage=3, deflate=True, clean=True)
        finally:
            pdf_document.close()

        # Serve the merged PDF file for download
        with open(merged_pdf_path, 'rb') as f:
            response = HttpResponse(f.read(), content_type='application/pdf')
            response['Content-Disposition'] = 'attachment; filename="merged.pdf"'

        # Remove the merged PDF file from the server
        os.remove(merged_pdf_path)

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
    pdf_reader = PyPDF2.PdfFileReader(input_file)
    pdf_writer = PyPDF2.PdfFileWriter()

    for page_num in range(len(pdf_reader.pages)):
        pdf_writer.add_page(pdf_reader.pages[page_num])

    pdf_writer.encrypt(password)

    # Create a temporary file for the encrypted PDF
    _, output_filename = tempfile.mkstemp(suffix='.pdf')
    with open(output_filename, 'wb') as output_file:
        pdf_writer.write(output_file)

    return output_filename


def PdfEncrypter(request):
    if request.method == 'POST' and 'file' in request.FILES:
        selected_file = request.FILES['file']
        password = request.POST['password']

        # Save the selected file temporarily on the server
        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
            temp_file.write(selected_file.read())
            temp_filename = temp_file.name

        # Secure the PDF file with the user-provided password
        encrypted_pdf_path = secure_pdf(temp_filename, password)

        # Prepare the encrypted PDF file for download
        with open(encrypted_pdf_path, 'rb') as f:
            encrypted_content = f.read()
            response = HttpResponse(encrypted_content, content_type='application/pdf')
            response['Content-Disposition'] = 'attachment; filename="secured_file.pdf"'

        # Close the file handle before removing the temporary files
        # os.remove(temp_filename)
        # os.remove(encrypted_pdf_path)

        return response

    return render(request, 'pdf_encrypter.html')