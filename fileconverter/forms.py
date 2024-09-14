from django import forms

class UploadFileForm(forms.Form):
    file = forms.FileField()

class ChangeRequestFileForm(forms.Form):
    file = forms.FileField()
