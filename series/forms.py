#-*- coding: utf-8 -*-
from django import forms
from series.models import Series


class TxtForm(forms.Form):
   serie_txt = forms.CharField(max_length = 1000,
                               required=True,
                               widget=forms.widgets.Textarea(
                                  attrs={
                                     "placeholder": "Ejemplo: 1AL100-200al300-500 AL 1000 - 1200 al 1500",
                                     "class": "textarea form-control mb-2 mt-2 is-medium",
                                  }
                               ),
                               label = ""
                               )

   class Meta:
      model = Series
      exclude = ("user",)