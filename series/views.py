import os
from django.http import FileResponse
from django.shortcuts import render

from djangoSeriesApp.settings import BASE_DIR
from series.forms import TxtForm
from series.models import Series


def index(request):
    data = None
    if request.method == "POST":
        form = TxtForm(request.POST)
        if form.is_valid():
            print(form.cleaned_data['serie_txt'])
            s = Series(series_text=form.cleaned_data['serie_txt'])
            s.escribirExcel()
            s.save()
            data = s
            data.action = s.series_text_result.split("\n")[len(s.series_text_result.split("\n")) - 1]
            data.action = data.action.replace(".xlsx", "")
            # data.series_text_result = data.series_text_result.replace('\n', "<br>")
    else:
        form = TxtForm()
    return render(request, "Series_idx.html", {"form": form, "data": data})



def down(request, num=0):

    file = BASE_DIR.__str__() + '/series/xlsx/series.' + str(num) + '.xlsx'

    path_to_file = os.path.realpath(file)
    response = FileResponse(open(path_to_file, 'rb'))
    file_name = 'series.' + str(num) + '.xlsx'
    response['Content-Disposition'] = 'inline; filename=' + file_name
    return response

