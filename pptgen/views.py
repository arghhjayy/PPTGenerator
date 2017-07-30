from django.shortcuts import render
# from django.http import HttpResponse
from pptgen.test import BASE_DIR, test

# Create your views here.

def main(request):
    if request.method == 'GET':
        return render(request, 'pptgen/main.html')
    else:
        ppt_topic = request.POST['ppt_topic']
        test(ppt_topic)
        PPT_DIR = BASE_DIR + '/pptgen/PPTS'
        return render(request, 'pptgen/generated.html', {'ppt_topic': ppt_topic, 'PPT_DIR': PPT_DIR})
