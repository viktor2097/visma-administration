[![PyPI version](https://badge.fury.io/py/visma-administration.svg)](https://badge.fury.io/py/visma-administration)
![PyPI - Downloads](https://img.shields.io/pypi/dm/visma-administration)
![GitHub](https://img.shields.io/github/license/viktor2097/visma-administration)
# *Python Visma Administration*  
Requires Python executable to be 32-bit to communicate with Visma API.  
```  
pip install visma-administration  
```  
  
# Quick functionality overview  
Easy integration with Visma Administration through pythonnet and Vismas Adk4NetWrapper.dll  

### Initialize communication with your database(s)
```py
from visma_administration import Visma
Visma.add_company(name="FTG10", common_path="Z:\\Gemensamma filer", company_path="Z:\\Företag\\FTG10")
```
### Get access to API for a company by calling Visma.get_company_api
```py
with Visma.get_company_api("FTG10") as api:
	pass
```
### Get a single record
```py
with Visma.get_company_api("FTG10") as api:
	record = api.supplier_invoice_head.get(adk_sup_inv_head_invoice_number="number")
	print(record.adk_sup_inv_head_invoice_number)
	record.adk_sup_inv_head_invoice_date = datetime.now() # supports date assignments  
	record.adk_sup_inv_head_vat_amount = 5000 # supports float and int 
	record.adk_sup_inv_head_paymstop = True # supports boolean assignments 
	record.adk_sup_inv_head_supplier_name = "Name" # supports string assignments 
	record.save()
```

### Get multiple records
```py
with Visma.get_company_api("FTG10") as api:
	invoices = api.supplier_invoice_head.filter(adk_sup_inv_head_invoice_number="f03*")
	total_sum = 0
	for invoice in invoices:
		total_sum += invoice.adk_sup_inv_head_total
	print(f"Total for all invoices with invoice_number starting as 'f03' is: {total_sum}")

	# filter returns a generator, so you can limit how many items to return
	import itertools  
	suppliers = api.supplier.filter(adk_supplier_name="*N*")  
	suppliers = itertools.islice(suppliers, 5)  
	for supplier in suppliers:  
		print(supplier.adk_supplier_name)  
```
### Create a record
```py
new_record = api.supplier.new() # .new() gives you a fresh object to work with
new_record.adk_supplier_name = "Nvidia"
###########################################################
# Calling .save() instead of .create() when creating an obj
# can result in some in-memory record being overwritten
new_record.create() # CALL THIS INSTEAD OF .SAVE()  
###########################################################
```

### Delete a record
```py
john = api.supplier.get(adk_supplier_name="John")
john.delete()
```

## Working with rows.
Existing rows can easily be accessed with **.rows()** and creating new ones by calling **.create_rows()**
```py
with Visma.get_company_api("FTG10") as api:
	# lets create a new invoice and add rows
	invoice = api.supplier_invoice_head.new()
	invoice.adk_sup_inv_head_invoice_number = "50294785"
	
	rows = invoice.create_rows(quantity=3) # rows is now a list containg 3 row objects!
	rows[0].adk_ooi_row_account_number = "5010"
	rows[1].adk_ooi_row_account_number = "4050"
	rows[2].adk_ooi_row_account_number = "3020"
	invoice.create()
	# Done. We now have a new invoice with 3 rows.

	# Here we can access existing rows
	for row in invoice.rows():
		print(row.adk_ooi_row_account_number)
	
	# Delete last row
	# If you would like to delete more than one row,
	# read the warnings further down on the readme 
	invoice.rows()[-1].delete()
	
	
```

# Deleting multiple rows
After **.delete()** is called on a row, Visma automatically reassigns new indexes to every row, therefore according to their documentation, you have to request all rows again to continue working with them.
```py
# If you would like to delete all custom rows, use the code snippet below.
# You have to start deleting them by negative index, since templates
# will automatically readd rows depending on your configurations
nrows = invoice.adk_sup_inv_head_nrows  
for i in range(int(nrows)):  
  invoice.rows()[-1].delete()
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
