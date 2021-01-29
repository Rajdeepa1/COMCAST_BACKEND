from django.urls import path, include
from django.conf.urls import url
from comcast_run_script.views import TestClass,PcnClass,TopicSearch

urlpatterns = [
    path('test', TestClass.as_view(), name='sample_test_case'),
    path('pcn', PcnClass.as_view(), name='sample_test_case2'),
    path('topic',TopicSearch.as_view(),name='topic_search')
    
]
