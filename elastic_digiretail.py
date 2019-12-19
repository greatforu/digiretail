import pandas
import sys
import json
from pprint import pprint
from elasticsearch import Elasticsearch

excel_data_df = pandas.read_excel('beginner_assignment01.xlsx', sheet_name='product_listing')

json_str = excel_data_df.to_json()
stat_file = open('elastic_data.json', 'w')
stat_file.write(json_str)

print('Excel Sheet to JSON:\n', json_str)

es = Elasticsearch(
    ['101.53.136.181'],
    port=9200

)

MyFile= open("/home/mayank/contextapi/elastic_data.json",'r').read()
ClearData = MyFile.splitlines(True)
i=0
json_str=""
docs ={}
for line in ClearData:
    line = ''.join(line.split())
    if line != "},":
        json_str = json_str+line
    else:
        docs[i]=json_str+"}"
        json_str=""
        print(docs[i])
        es.index(index='Digiretail', doc_type='Data', id=i, body=docs[i])
        i=i+1