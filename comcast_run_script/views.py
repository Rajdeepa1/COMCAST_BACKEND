from django.shortcuts import render
import subprocess
from rest_framework import status
from rest_framework.response import Response
from rest_framework.views import APIView
from django.http import HttpResponse
from comcast_run_script.scripts.test1 import func
from comcast_run_script.scripts.pcnfn import func_pcn
from comcast_run_script.scripts.topicsearch import func_topic
import json
from django.http import StreamingHttpResponse
def convert(lst): 
    return eval(lst) 
# Create your views here.
class TestClass(APIView):
    def get(self, request):
        #useless_cat_call = subprocess.run(["python","C:/Users/rajdchak/Documents/DEVELOPMENT/Comcast_site/comcast/comcast_run_script/test1.py"], stdout=subprocess.PIPE, text=True, input="Hello from the other side")
        #print(useless_cat_call.stdout)
        #l=['CSCvf69272','CSCvg54149','CSCvj78551','CSCvk00895','CSCvm07353']
        l=request.GET["p"].split(',')
        print(l)
        print(type(l))
        
        file_name="bugdets.xlsx"
        output=func(l)
        response = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename=%s' % file_name
        return response
        #  vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv             
class PcnClass(APIView):
    def get(self, request):
        
        
        file_name="pcn.pptx"
        l=request.GET["p"].split(',')
        print(l)
        print(type(l))
        output=func_pcn(l)
        response = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        response['Content-Disposition'] = 'attachment; filename=%s' % file_name
        return response  

class TopicSearch(APIView):
    def get(self, request):
        #useless_cat_call = subprocess.run(["python","C:/Users/rajdchak/Documents/DEVELOPMENT/Comcast_site/comcast/comcast_run_script/test1.py"], stdout=subprocess.PIPE, text=True, input="Hello from the other side")
        #print(useless_cat_call.stdout)
        #l=['CSCvf69272','CSCvg54149','CSCvj78551','CSCvk00895','CSCvm07353']
        #l=request.GET["p"].split(',')
        #print(l)
        #print(type(l))
        l=request.GET["p"]
        
        file_name="Service_requests.xlsx"
        output=func_topic(l)
        response = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename=%s' % file_name
        return response


                                                           

