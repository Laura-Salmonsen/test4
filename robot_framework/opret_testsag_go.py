import requests
import os
import json
from requests_ntlm import HttpNtlmAuth
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import process_laura as p




orchestrator_connection = OrchestratorConnection("GOTest", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
API_url = orchestrator_connection.get_constant("GOApiTESTURL").value
go_credentials = orchestrator_connection.get_credential("GOTestApiUser")
session = requests.Session()
session.auth = HttpNtlmAuth(go_credentials.username, go_credentials.password)
session.post(API_url, timeout=500)

url = "https://testad.go.aarhuskommune.dk/geosager/_goapi/Cases"

payload = json.dumps({
  "CaseTypePrefix": "GEO",
  "MetadataXml": "<z:row xmlns:z=\"#RowsetSchema\" ows_Title=\"Case From Api Test Laura3\" ows_CaseStatus=\"Åben\" ows_EksterntSagsID=\"TestSagID\" ows_EksterntSystemID=\"TestSystemID\" ows_EksternLink=\"https://testad.go.aarhuskommune.dk/cases/GEO24/GEO-2021-000088/SitePages/Home.aspx\" />",
  "ReturnWhenCaseFullyCreated": True
})
headers = {
  'Content-Type': 'application/json'
}

response = session.post(url, headers=headers, data=payload, timeout = 500)

print(response.text)