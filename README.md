# *Python Visma Administration*  
Requires Python executable to be 32-bit to communicate with Visma API.  
```  
pip install visma-administration  
```  
  
# Quick functionality overview  
Seamless integration with Visma Administration through pythonnet and Vismas Adk4NetWrapper.dll  
```py  
from visma_administration import VismaAPI  
from datetime import datetime  
  
# Initilize a connection with one or multiple companies  
company_1 = VismaAPI(common_path="Z:\\Gemensamma filer", company_path="Z:\\Företag\\FTG9")  
company_2 = VismaAPI(common_path="Z:\\Gemensamma filer", company_path="Z:\\Företag\\FTG10")  
  
# Retrieve a single supplier  
john = company_1.supplier.get(adk_supplier_name="John")  
  
# Print data about john  
print(john.adk_supplier_short_name)  
print(john.adk_supplier_credit_limit)  
  
# Assign new data to john  
john.adk_supplier_short_name = "JN"             # supports string assignments  
john.adk_supplier_autogiro = True               # supports boolean assignments  
john.adk_supplier_credit_limit = 50000          # supports float and int assignments  
john.adk_supplier_timestamp = datetime.now()    # supports date assignments  
john.save() # save to the database  

# Delete a record
john.delete()
  
# Retrieve all suppliers that contain DE anywhere in its name - returns a generator  
for supplier in company_2.supplier.filter(adk_supplier_name="*DE*"):  
 print(supplier.adk_supplier_name)
 
# Create a new record  
new_record = Supplier().new()  
new_record.adk_supplier_name = "Nvidia"  
new_record.create() # CALL THIS INSTEAD OF .SAVE()  
  
# Get 5 elements  
import itertools  
suppliers = company_1.supplier.filter(ADK_SUPPLIER_NAME="*N*")  
suppliers = itertools.islice(suppliers, 5)  
for supplier in suppliers:  
 print(supplier.adk_supplier_name)  
```  
  
# How it works  
  
The project tries to reduce a significant amount of boilerplate code to get working with Visma Administration.  
  
It makes boilerplate like this  
  
```py  
... import C# libraries with pythonnet  
  
Api.AdkOpen2(common_path, company_path, username, password)  
john = Api.AdkCreateData(visma.api.ADK_DB_SUPPLIER)  
Api.AdkSetFilter(john, 1, "John", 0)  
Api.AdkFirstEx(john, False)  
from System import String  
print(Api.AdkGetStr(john, Api.ADK_SUPPLIER_NAME, String("")))  
Api.AdkSetStr(john, Api.ADK_SUPPLIER_NAME, String("NewName")  
Api.AdkUpdate(john)  
```  
  
Turn into  
  
```py  
from visma_administration import VismaAPI  
  
visma = VismaAPI(common_path, company_path)  
john = visma.supplier.get(adk_supplier_name="John")  
print(john.adk_supplier_name)  
john.adk_supplier_name = "New John"  
john.save()  
```  
  
In Visma's documentation you can find a list of DB_fields that you may access through their API.  
For instance  
```  
ADK_DB_PROJECT  
ADK_DB_ACCOUNT  
ADK_DB_SUPPLIER 
```  
  
You can access these fields directly from your instantiated VismaAPI object,   
ADK_DB_ is removed and the name is lowercased.  
  
For example:  
      
```py  
visma = VismaAPI(common_path, company_path)  
visma.supplier  
visma.account  
visma.project  
```  
  
These DB_Field attributes returns an object from which you can request data  
  
```py  
visma.supplier.get()  # returns a single object or returns an error  
visma.supplier.filter()  # returns multiple objects  
visma.supplier.new()  # Gives you a new record  
```

Both `get` and `filter` accepts filtering on a field as argument.
You pass the field name which you want to filter upon, and a valid filter expression ( Visma have documentation for this )

```py
"""
This filters on the ADK_SUPPLIER_NAME field of ADK_DB_SUPPLIER
*text* is a valid filter expression which returns a result with inc anywhere inside of the name
Refer to Visma documentation for further information on how to form different types of filter expressions.
"""
result = visma.supplier.get(adk_supplier_name="*inc*")
```

Consider whatever `get`, `filter` and `new` returns to be database record(s).
These records have fields associated with them, if you check documentation for ADK_DB_SUPPLIER, you can find a whole bunch of fields associated with it.
```
ADK_SUPPLIER_FIRST
ADK_SUPPLIER_NUMBER
ADK_SUPPLIER_NAME
ADK_SUPPLIER_SHORT_NAME
... And a lot more
```
These fields are mainly used in two ways,
```py
test = visma.supplier.get(adk_supplier_name="test")  # Used when filtering
test.adk_supplier_name = "test1"  # Or used when interacting with a supplier record
```
 
  
## Additional information
  
[Documentation for the api can be found here](https://vismaspcs.se/support/utvecklarpaket-eget-bruk)  
  
*Visma has to be installed on the PC you import the package on. Either full client or integration client is fine.* 

  
## License  
  
[![License](http://img.shields.io/:license-mit-blue.svg?style=flat-square)](http://badges.mit-license.org)  
  
- **[MIT license](http://opensource.org/licenses/mit-license.php)**