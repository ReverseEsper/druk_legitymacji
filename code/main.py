
from druk_legitymacji import get_from_xls as gfx
from druk_legitymacji import work_on_excel as woe


uczniowie = gfx('input_file.xlsx')
woe('template.xls','plik_legitymacji.xls',uczniowie)
#def work_on_excel(template_path,output_path,input_data):

