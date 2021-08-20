from django import forms

class FileForm(forms.Form):

    input_filee = forms.FileField(label="Submit file")
