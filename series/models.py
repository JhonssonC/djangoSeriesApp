import xlsxwriter
from django.db import models

# Create your models here.
from django.db import models

from djangoSeriesApp.settings import BASE_DIR


class Series(models.Model):
    series_text = models.CharField(max_length=1000)
    series_arr_result = models.CharField(max_length=1000)
    series_text_result = models.CharField(max_length=1000)
    req_date = models.DateTimeField('Fecha de Requerimiento', auto_now=True)

    def __str__(self):
        return self.series_arr_result

    def escribirExcel(self):

        series = self.series_text
        self.series_text_result=""

        series = series.replace('\n', '')
        series = series.replace(' ', '')
        series = series.replace('al', 'AL')

        if self.indexof(series, '-') > 0:
            series = series.split('-')
        else:
            series = [series]
        try:
            nx = str(Series.objects.count())
        except Exception:
            nx=str(0)

        workbook = xlsxwriter.Workbook(BASE_DIR.__str__()+'/series/xlsx/series.'+nx+'.xlsx')
        worksheet = workbook.add_worksheet()
        valores_digitar_excel = []
        reg = 0
        dup = 'No'
        print(f'Encontradas {len(series)} series')
        self.series_text_result += f'\nEncontradas {len(series)} series'
        for serie in series:
            if serie.count('AL') == 1:
                al = self.indexof(serie, 'AL')
                if 0 < al < len(serie) - 1:
                    delimitadores = serie.split('AL')
                    try:
                        i = int(delimitadores[0])
                        j = int(delimitadores[1]) + 1
                        if i < j:
                            for l in range(i, j):
                                if self.indexof(valores_digitar_excel, l) == -1:
                                    valores_digitar_excel.append(l)
                                else:
                                    dup = 'Si'
                            print(f'Serie {str(series.index(serie) + 1).center(20)}  |  {serie.center(35)}   |   {str(j - i).center(10)} valores')
                            self.series_text_result += f'\nSerie {str(series.index(serie) + 1).center(20)}  |  {serie.center(35)}   |   {str(j - i).center(10)} valores'
                            reg += j - i
                        else:
                            print(f'Error en serie {str(series.index(serie) + 1).center(11)}  |  {serie.center(35)}   |   {str(0).center(10)} valores')
                            self.series_text_result += f'\nError en serie {str(series.index(serie) + 1).center(11)}  |  {serie.center(35)}   |   {str(0).center(10)} valores'
                    except Exception:
                        print(f'Error en serie {str(series.index(serie) + 1).center(11)}  |  {serie.center(35)}   |   {str(0).center(10)} valores')
                        self.series_text_result += f'\nError en serie {str(series.index(serie) + 1).center(11)}  |  {serie.center(35)}   |   {str(0).center(10)} valores'
                else:
                    print(f'Error en serie {str(series.index(serie) + 1).center(11)}  |  {serie.center(35)}   |   {str(0).center(10)} valores')
                    self.series_text_result += f'\nError en serie {str(series.index(serie) + 1).center(11)}  |  {serie.center(35)}   |   {str(0).center(10)} valores'
            else:
                try:
                    if self.indexof(valores_digitar_excel, int(serie)) == -1:
                        valores_digitar_excel.append(int(serie))
                    else: dup='Si'
                    print(f'Serie {str(series.index(serie) + 1).center(20)}  |  {serie.center(35)}   |   {str(1).center(10)} valor')
                    self.series_text_result += f'\nSerie {str(series.index(serie) + 1).center(20)}  |  {serie.center(35)}   |   {str(1).center(10)} valor'
                    reg += 1
                except Exception:
                    print(f'Error en serie {str(series.index(serie) + 1).center(11)}  |  {serie.center(35)}   |   {str(0).center(10)} valores')
                    self.series_text_result += f'\nError en serie {str(series.index(serie) + 1).center(11)}  |  {serie.center(35)}   |   {str(0).center(10)} valores'

        valores_digitar_excel.sort()

        worksheet.write('A1', 'ITEM')
        worksheet.write('B1', 'NUMERACION')
        for mynumber in valores_digitar_excel:
            celda_item = valores_digitar_excel.index(mynumber) + 1
            worksheet.write(celda_item, 0, celda_item)
            worksheet.write(celda_item, 1, mynumber)

        self.series_arr_result = str(valores_digitar_excel)

        workbook.close()

        print(f'Cantidad de Numeraciones subidas a Excel: {len(valores_digitar_excel)}')
        self.series_text_result += f'\nCantidad de Numeraciones subidas a Excel: {len(valores_digitar_excel)}'
        print(f'Cantidad de numeraciones digitadas: {reg}')
        self.series_text_result += f'\nCantidad de numeraciones digitadas: {reg}'

        print(f'Existen rangos con duplicidades: {dup}')
        self.series_text_result += f'\nExisten rangos con duplicidades: {dup}'

        print(BASE_DIR.__str__()+'/series/xlsx/series.'+str(Series.objects.count())+'.xlsx')
        self.series_text_result += '\nseries.'+str(Series.objects.count())+'.xlsx'


    def indexof(self, valor, encontrar):
        try:
            return valor.index(encontrar)
        except ValueError:
            return -1