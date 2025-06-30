from django import forms
from datetime import date

class MonthYearForm(forms.Form):
    _selected_action = forms.CharField(widget=forms.MultipleHiddenInput)
    month = forms.IntegerField(label="Month", min_value=1, max_value=12, initial=date.today().month)
    year = forms.IntegerField(label="Year", min_value=2020, max_value=2100, initial=date.today().year)
