from django.shortcuts import render
from django.views import View
from . import preprocessing
from .forms import FileForm
from django.core.files.storage import FileSystemStorage
# Import mimetypes module
import mimetypes
# import os module
from django.contrib import messages
import os
from django.http.response import HttpResponse
from django.http import FileResponse

def handle_uploaded_file(f):
    with open('/uploads/'+f.name, 'wb+') as destination:
        for chunk in f.chunks():
            destination.write(chunk)

def download_file(request):
    # Define Django project base directory
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    # Define text file name
    filename = 'Output.pptx'
    # Define the full file abspath
    if output_type=="ppt":
        filepath = r'C:\Users\lenovo\Desktop\FYP\backend\Output Presentations\OUTPUT.pptx'
        print("ppt")
    if output_type=="pdf":
        filepath = r'C:\Users\lenovo\Desktop\FYP\backend\Output Presentations\OUTPUT.pdf'
        print("pdf")
    # Open the file for reading content
    path = open(filepath, 'r')
    print("reaching")
    # Set the mime type
    mime_type, _ = mimetypes.guess_type(filepath)
    # Set the return value of the HttpResponse
    response = FileResponse(open(filepath, 'rb'))
    return response



class index(View):


    def get(self,request,*args,**kwargs):
        the_form=FileForm()
        context = {
        "title" : "Submit the file",
        "form"  : "the_form"
        }
        return render(request, 'orange.html',{'form':the_form})

    def post(self,request,*args,**kwargs):
        print(request.POST)
        input_format=request.POST.get('filetype')
        input_file=request.POST.get('inputfile')
        no_of_sentences=request.POST.get('noofslides')
        output_format=request.POST.get('optype')
        theme=request.POST.get('presentation')
        print(input_format)
        print(no_of_sentences)
        print(output_format)
        global output_type
        output_type=output_format
        print("--"+output_type)
        the_form=FileForm(request.POST,request.FILES)
        if the_form.is_valid():

            request_file = request.FILES['input_filee']
            if request_file:
            # save attatched file
            # create a new instance of FileSystemStorage
                fs = FileSystemStorage(r"C:\Users\lenovo\Desktop\FYP\backend\uploads")
                file = fs.save(request_file.name, request_file)
                fileurl = fs.url(file)
        print("preprocessing called")
        input_file=request_file.name
        preprocessing.preprocess(input_format,output_format,no_of_sentences,input_file,theme)

        messages.success(request, 'File Submitted successfully.')
        return render(request, 'orange.html')
