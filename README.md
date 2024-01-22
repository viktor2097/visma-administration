
[![PyPI version](https://badge.fury.io/py/visma-administration.svg)](https://badge.fury.io/py/visma-administration)
![PyPI - Downloads](https://img.shields.io/pypi/dm/visma-administration)
![GitHub](https://img.shields.io/github/license/viktor2097/visma-administration)
# *Python Visma Administration*  
API for Visma Administration 200/500/1000/2000. 

Visma has added support for a 64-bit API in version 2023.1. Therefore, this package aims to be compatible with both 32-bit and 64-bit versions to maintain backward compatibility. The AdkNet4Wrapper is loaded on 32-bit Python installations, and AdkNet6Wrapper is loaded on 64-bit Python installations.

Writing integrations for Visma's API often involves extensive boilerplate code. This package aims to present the API with a user-friendly, Pythonic interface, and make it more accessible. It's not feature complete, but it covers most basic needs, and you can tap into the underlying C# API if needed.
# Installation
```
pip install visma-administration  
```  

## Usage
### Add companies that you would like to access data from
```py
from visma_administration import Visma
Visma.add_company(name="FTG10", common_path="Z:\\Gemensamma filer", company_path="Z:\\Företag\\FTG10")

# You can then access specific companies with a context manager
with Visma.get_company_api("FTG10") as api:
	pass
```

## CRUD Operations
**Read a single record**
```py
with Visma.get_company_api("FTG10") as api:
    record = api.supplier_invoice_head.get(adk_sup_inv_head_invoice_number="number")
    print(record.adk_sup_inv_head_invoice_number)
```
**Update a record**
```py
with Visma.get_company_api("FTG10") as api:
    record = api.supplier_invoice_head.get(adk_sup_inv_head_invoice_number="number")
    record.adk_sup_inv_head_invoice_date = datetime.now()
    record.adk_sup_inv_head_vat_amount = 5000
    record.adk_sup_inv_head_paymstop = True
    record.adk_sup_inv_head_supplier_name = "Name"
    record.save()
```

**Create a record**
```py
with Visma.get_company_api("FTG10") as api:
    new_record = api.supplier.new() # new method is called
    new_record.adk_supplier_name = "Nvidia"
    new_record.save()
```

**Delete a record**
```py
with Visma.get_company_api("FTG10") as api:
    john = api.supplier.get(adk_supplier_name="John")
    john.delete()
```

**Get multiple records**
```py
with Visma.get_company_api("FTG10") as api:
	invoices = api.supplier_invoice_head.filter(adk_sup_inv_head_invoice_number="f03*")
	total_sum = sum(invoice.adk_sup_inv_head_total for invoice in invoices)
	print(f"Total for all invoices with invoice_number starting as 'f03' is: {total_sum}")

	# filter returns a generator, so you can limit how many items to return
	import itertools  
	suppliers = api.supplier.filter(adk_supplier_name="*N*")  
	suppliers = itertools.islice(suppliers, 5)  
	for supplier in suppliers:  
		print(supplier.adk_supplier_name)  
```
**Working with rows**
 Existing rows can be accessed with: **.rows()** and creating new ones by calling **.create_rows()**
```py
with Visma.get_company_api("FTG10") as api:
	# lets create a new invoice and add rows
	invoice = api.supplier_invoice_head.new()
	invoice.adk_sup_inv_head_invoice_number = "50294785"
	
	# Creates 3 row objects
	rows = invoice.create_rows(quantity=3)
	rows[0].adk_ooi_row_account_number = "5010"
	rows[1].adk_ooi_row_account_number = "4050"
	rows[2].adk_ooi_row_account_number = "3020"
	invoice.save()
	# We now have a new invoice with 3 rows.

	# Here we can access existing rows
	for row in invoice.rows():
		print(row.adk_ooi_row_account_number)
	
	# Delete last row
	# If you would like to delete more than one row,
	# Check the Deleting multiple rows section of the README
	invoice.rows()[-1].delete()
	
	
	
```

### Deleting multiple rows
After **.delete()** is called on a row, Visma automatically reassigns new indexes to every row, therefore according to their documentation, you have to request all rows again to continue working with them.
```py
# If you would like to delete all custom rows, use the code snippet below.
# You have to start deleting them by negative index, since templates
# will automatically readd rows depending on your configurations
nrows = invoice.adk_sup_inv_head_nrows  
for i in range(int(nrows)):  
  invoice.rows()[-1].delete()
```

# Other info
In Visma's documentation you can find a list of DB_fields that you may access through their API.  
For instance  
```  
ADK_DB_PROJECT  
ADK_DB_ACCOUNT  
ADK_DB_SUPPLIER 
```  
  
You can access these fields directly from your instantiated Visma API object,   
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
You pass the field name which you want to filter upon, and a valid filter expression ( Visma have documentation for this.
You can filter on multiple fields by just passing multiple filter expressions

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
