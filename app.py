import json
import pandas as pd
import openpyxl
from tkinter import Tk

# data = input("Enter Json String")

data = "{\"ListTitle\":null,\"FeaturedCompanies\":null,\"Companies\":[{\"CompanyName\":\"ΚΥΡΙΑΚΟΠΟΥΛΟΣ ΒΑΣΙΛΕΙΟΣ\",\"CompanyNameEnglish\":\"KYRIAKOPOULOS BASIL\",\"CompanyId\":601001307575,\"LogoLink\":null,\"BusinessUnit\":60,\"PrefectureCode\":138,\"PrefectureName\":\"Αττική\",\"PrefectureNameEnglish\":\"Attiki\",\"Sales\":null,\"NumberOfEmployees\":null,\"Nace2Description\":\"1712-Κατασκευή χαρτιού και χαρτονιού\",\"Nace2level1\":\"17\",\"MainNace2\":\"1712\",\"LeadScore\":null,\"isSelected\":false,\"FriendlyUrl\":\"/el/company/kyriakopoulos-basil/601001307575\",\"Base64Img\":null,\"DateAdded\":null,\"RemovedFromFBZdata\":false,\"Longitude\":23.72535650000000000000,\"Latitude\":37.97905410000000000000,\"LegalForm\":\"Sole Proprietorship\",\"VatNumber\":\"BBD240F734D19352ED43275D220948560BC4B5864186B07A8B2C084D3318FFB0\",\"IcapSector\":[\"63\"]},{\"CompanyName\":\"ATTICA SOFT PAPER Α.Β.Ε.Ε.\",\"CompanyNameEnglish\":\"ATTICA SOFT PAPER INDUSTRIAL AND COMMERCIAL S.A\",\"CompanyId\":601000370498,\"LogoLink\":null,\"BusinessUnit\":60,\"PrefectureCode\":138,\"PrefectureName\":\"Αττική\",\"PrefectureNameEnglish\":\"Attiki\",\"Sales\":\"0.0\",\"NumberOfEmployees\":null,\"Nace2Description\":\"1712-Κατασκευή χαρτιού και χαρτονιού\",\"Nace2level1\":\"17\",\"MainNace2\":\"1712\",\"LeadScore\":null,\"isSelected\":false,\"FriendlyUrl\":\"/el/company/attica-soft-paper-industrial-and-commercial-sa/601000370498\",\"Base64Img\":null,\"DateAdded\":null,\"RemovedFromFBZdata\":false,\"Longitude\":23.72671880000000000000,\"Latitude\":37.96712740000000000000,\"LegalForm\":\"Societe Anonyme\",\"VatNumber\":\"2E3AA5EE86AFFB2BEF3009656FD524DE8EA1391C08F4371EBE7A042C68997894\",\"IcapSector\":[\"63\"]}],\"Pagination\":{\"ItemsPerPage\":10,\"TotalItems\":2,\"CurrentPage\":1,\"FirstPage\":1,\"NextPage\":2,\"PreviousPage\":0,\"LastPage\":1,\"TotalPages\":1,\"PreviousResultsCount\":0,\"ForwardResultsCount\":2,\"CurrentPageResultsCount\":2,\"NearPages\":[1],\"HasMultiplePages\":false},\"JsonRequest\":null,\"DfpNaceTargets\":null,\"DfpIcapSectorTargets\":null,\"DynamicListCriteria\":null}"

dict = json.loads(data)
name = []
secondValue= []
place = []
url = []

x = dict.get('Companies')


for items in range(len(x)):
		name.append(x[items].get('CompanyName'))
		secondValue.append(x[items].get('Nace2Description'))
		place.append(x[items].get('PrefectureName'))
		url.append("www.findbiz.gr" + x[items].get('FriendlyUrl'))

df = pd.DataFrame({
    'ΕΠΩΝΥΜΙΑ': name,
    'ΔΡΑΣΤΗΡΙΟΤΗΤΑ': secondValue,
    'ΠΕΡΙΟΧΗ': place,
    'URL' : url,
}) 

name = input("Give excel name")
 
# Write DataFrame to Excel 
df.to_excel(name+'.xlsx')
