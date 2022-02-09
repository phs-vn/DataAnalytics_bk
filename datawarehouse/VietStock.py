from request import *
import zeep
from zeep import Client



URL = 'http://s1.vietstock.vn/DataWS/DataWebService.svc?wsdl'


client = Client(wsdl=URL)
result = client.service.GetData(
    _GroupType = 4,
    _DataType = 6,
    _FromDate = dt.datetime(2021,12,2),
    _ToDate =  dt.datetime(2022,1,3),
    )
a = zeep.helpers.serialize_object(result)
print(a)



outer_dicts = a['_value_1'][1]['_value_1']
tables = (pd.DataFrame(outer_dicts[row]['Table'],index=[row]) for row in range(len(outer_dicts)))

table = pd.concat(tables)



table.to_excel('vietstock.xlsx')