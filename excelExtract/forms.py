from django import forms
from excelExtract.models import document


class uploadDocumentForm(forms.ModelForm):
    class Meta:
        model = document
        fields =  ('document',)