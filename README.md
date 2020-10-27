# *Python Visma Administration*  
Requires Python executable to be 32-bit to communicate with Visma API.  
```  
pip install visma-administration  
```  
  
# Quick functionality overview  
Easy integration with Visma Administration through pythonnet and Vismas Adk4NetWrapper.dll  
```py  
from visma_administration import Visma  
from datetime import datetime  
  
# Add companies you would like to access
# It's highly recommended to only use one company, which i'll explain why further down in the readme
Visma.add_company(name="FTG9", common_path="Z:\\Gemensamma filer", company_path="Z:\\Företag\\FTG9")  
Visma.add_company(name="FTG10", common_path="Z:\\Gemensamma filer", company_path="Z:\\Företag\\FTG10")  

# Gets API for FTG10 we added earlier
with Visma.get_company_api("FTG10") as api:
    # Retrieve a single supplier  
    john = api.supplier.get(adk_supplier_name="John")  
      
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
    for supplier in api.supplier.filter(adk_supplier_name="*DE*"):  
     print(supplier.adk_supplier_name)
     
    # Create a new record  
    new_record = api.supplier.new()  
    new_record.adk_supplier_name = "Nvidia"  
    new_record.create() # CALL THIS INSTEAD OF .SAVE()  
      
    # Get 5 elements  
    import itertools  
    suppliers = api.supplier.filter(ADK_SUPPLIER_NAME="*N*")  
    suppliers = itertools.islice(suppliers, 5)  
    for supplier in suppliers:  
     print(supplier.adk_supplier_name)  
```

# Why is it recommended to only configure one company?

The underlying API is a DLL file, which can only be loaded once per process, and the fact that Vismas API has no direct way to open 
multiple companies at once.
I had to choose between complaticating the codebase a lot and try to load the dll in multiple processes for each company, or simply allow one company to be accessed at a time.

The latter works well for me, as we don't have too much requests. Whenever your multithreaded code tries to open two different companies, it will finish all open API's through the context manager (get_company_api) before opening the other company.
If you for instance have 100 managers open for company A, you need to wait for those to finish before the context manager yields the API for company B. This could be problematic if you make a ton of requests, and could add some major delay.

Therefore it's up to you if you want to configure multiple companies or not.

# Other info
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
with visma.get_company_api("company") as api:
    api.supplier  
    api.account  
    api.project  
```  
  
These DB_Field attributes returns an object from which you can request data  
  
```py  
api.supplier.get()  # returns a single object or returns an error  
api.supplier.filter()  # returns multiple objects  
api.supplier.new()  # Gives you a new record  
```

Both `get` and `filter` accepts filtering on a field as argument.
You pass the field name which you want to filter upon, and a valid filter expression ( Visma have documentation for this )

```py
"""
This filters on the ADK_SUPPLIER_NAME field of ADK_DB_SUPPLIER
*text* is a valid filter expression which returns a result with inc anywhere inside of the name
Refer to Visma documentation for further information on how to form different types of filter expressions.
"""
result = api.supplier.get(adk_supplier_name="*inc*")
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
test = api.supplier.get(adk_supplier_name="test")  # Used when filtering
test.adk_supplier_name = "test1"  # Or used when interacting with a supplier record
```
 
  
## Additional information
  
[Documentation for the api can be found here](https://vismaspcs.se/support/utvecklarpaket-eget-bruk)  
  
*Visma has to be installed on the PC you import the package on. Either full client or integration client is fine.* 

  
## License  
  
[![License](http://img.shields.io/:license-mit-blue.svg?style=flat-square)](http://badges.mit-license.org)  
  
- **[MIT license](http://opensource.org/licenses/mit-license.php)**