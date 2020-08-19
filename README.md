# *What is Visma-administration?*
```
pip install visma-administration
```

It's a simple wrapper around the *AdkNet4Wrapper.dll*

It lets you execute code from the .NET wrapper for Visma Administration 200/500/1000/2000, directly from Python.

```py
from visma_administration import VismaAPI

# common_path and company_path are documented in the AdkOpen(CHAR* pszSystemPath, CHAR* pszFtgPath) function
visma = VismaAPI(common_path="Z:\\Gemensamma filer", company_path="Z:\\Företag\\FTG9")

# Call functions like any C# code!
pdata = visma.api.AdkCreateData(visma.api.ADK_DB_SUPPLIER)
visma.api.AdkFirstEx(pdata, False)

# AdkGetStr() takes a String out parameter, so we can import String from the System library,
# and it will be saved to the variable supplier_name instead
from System import String
supplier_name = visma.api.AdkGetStr(pdata, visma.api.ADK_SUPPLIER_NAME, String(""))

```

However, this is cumbersome, it's possible to use a simpler interface

Classes for each `ADK_DB_FIELDNAME` is created dynamically (you can find them below in this README or more accurately from Adk.h from the documentation)

Classes follow a naming convention of Fieldname
```py
# For example, ADK_DB_SUPPLIER is one class that is created dynamically and ready to be imported
# Note that the class name is generated as a title version of the field name
from visma_administration import VismaAPI, Supplier
visma = VismaAPI(common_path="Z:\\Gemensamma filer", company_path="Z:\\Företag\\FTG9")
```

ADK_DB_FILEDNAME have other fields associated with it! Again, you can find the entire list in Adk.h

As an example, ADK_DB_SUPPLIER has among others these fields associated with it
```c
#define ADK_SUPPLIER_NAME                                        1              //eChar 
#define ADK_SUPPLIER_SHORT_NAME                                  2              //eChar 
```
## .get()
We can retrieve a single record with the .get() method of a DB class
```py
# We pass the field name to filter on, get retrieves the first record found or raises an exception if not able to get a record
record = Supplier().get(ADK_SUPPLIER_NAME="*hello*")

# You can now access fields like any attributes! (lowercased)
print(record.adk_supplier_name) # output: hello world
print(record.adk_supplier_short_name) # output: hw

# You can even assign new data!
record.adk_supplier_name = "hello world!"
record.save() # call save to update the record to the database
print(record.adk_supplier_name) # output: hello world!

# You can delete the record too
record.delete()
```

## .filter()
filter is a python generator that returns multiple records

```py
# Updates multiple records
for record in Supplier().filter(ADK_SUPPLIER_NAME="*hello*"):
    record.adk_supplier_name = SOME_VALUE
    record.save()
```

## Filtering

To filter with **.get()** and **.filter()**, provide the field name to filter on and assign a filter expression to it.
You can find documentation on how to construct filter expressions in the documentation
```py
Supplier().get(ADK_SUPPLIER_NAME="filter expression")
```

## Requirements

[Documentation for the wrapper can be found here](https://vismaspcs.se/support/utvecklarpaket-eget-bruk)

**Visma has to be installed on the PC you import the package on. Either full client or integration client is fine.**

**visma-administration will automatically find the installation location of the api on your PC based on environment variables visma installation adds.**

## ADK_DB_FIELDS
```
#define ADK_DB_CUSTOMER                                          0
#define ADK_DB_ARTICLE                                           1
#define ADK_DB_ORDER_HEAD                                        2
#define ADK_DB_ORDER_ROW                                         3
#define ADK_DB_OFFER_HEAD                                        4
#define ADK_DB_OFFER_ROW                                         5
#define ADK_DB_INVOICE_HEAD                                      6
#define ADK_DB_INVOICE_ROW                                       7
#define ADK_DB_SUPPLIER_INVOICE_HEAD                             8
#define ADK_DB_SUPPLIER_INVOICE_ROW                              9
#define ADK_DB_PROJECT                                          10
#define ADK_DB_ACCOUNT                                          11
#define ADK_DB_SUPPLIER                                         12
#define ADK_DB_CODE_OF_TERMS_OF_DELIVERY                        13
#define ADK_DB_CODE_OF_WAY_OF_DELIVERY                          14
#define ADK_DB_CODE_OF_TERMS_OF_PAYMENT                         15
#define ADK_DB_CODE_OF_LANGUAGE                                 16
#define ADK_DB_CODE_OF_CURRENCY                                 17
#define ADK_DB_CODE_OF_CUSTOMER_CATEGORY                        18
#define ADK_DB_CODE_OF_DISTRICT                                 19
#define ADK_DB_CODE_OF_SELLER                                   20
#define ADK_DB_DISCOUNT_AGREEMENT                               21
#define ADK_DB_CODE_OF_ARTICLE_GROUP                            22
#define ADK_DB_CODE_OF_ARTICLE_ACCOUNT                          23
#define ADK_DB_CODE_OF_UNIT                                     24
#define ADK_DB_CODE_OF_PROFIT_CENTRE                            25
#define ADK_DB_CODE_OF_PRICE_LIST                               26
#define ADK_DB_PRM                                              27
#define ADK_DB_INVENTORY_ARTICLE                                28
#define ADK_DB_MANUAL_DELIVERY_IN                               29
#define ADK_DB_MANUAL_DELIVERY_OUT                              30
#define ADK_DB_DISPATCHER										31
#define ADK_DB_BOOKING_HEAD										32
#define ADK_DB_BOOKING_ROW										33
#define ADK_DB_CODE_OF_CUSTOMER_DISCOUNT_ROW					34
#define ADK_DB_CODE_OF_ARTICLE_PARCEL							35
#define ADK_DB_CODE_OF_ARTICLE_NAME								36
#define ADK_DB_PRICE											37
#define ADK_DB_ARTICLE_PURCHASE_PRICE							38
#define ADK_DB_CODE_OF_WAY_OF_PAYMENT							39
#define ADK_DB_FREE_CATEGORY_1									40 //SPCS F�rening
#define ADK_DB_FREE_CATEGORY_2									41 //SPCS F�rening
#define ADK_DB_FREE_CATEGORY_3									42 //SPCS F�rening
#define ADK_DB_FREE_CATEGORY_4									43 //SPCS F�rening
#define ADK_DB_FREE_CATEGORY_5									44 //SPCS F�rening
#define ADK_DB_FREE_CATEGORY_6									45 //SPCS F�rening
#define ADK_DB_FREE_CATEGORY_7									46 //SPCS F�rening
#define ADK_DB_FREE_CATEGORY_8									47 //SPCS F�rening
#define ADK_DB_FREE_CATEGORY_9									48 //SPCS F�rening
#define ADK_DB_FREE_CATEGORY_10									49 //SPCS F�rening
#define ADK_DB_MEMBER                                           50 //SPCS F�rening
#define ADK_DB_DELIVERY_NOTE_HEAD								51
#define ADK_DB_DELIVERY_NOTE_ROW								52
#define ADK_DB_PACKAGE_HEAD                                     53
#define ADK_DB_PACKAGE_ROW                                      54
#define ADK_DB_IMP_PACKAGE_HEAD                                 55
#define ADK_DB_IMP_PACKAGE_ROW                                  56
#define ADK_DB_DELIVERY_ADDRESS                                 57
#define ADK_DB_PRM2												58
#define ADK_DB_CODE_OF_YOUR_REF_CUSTOMER						59
#define ADK_DB_CODE_OF_YOUR_REF_SUPPLIER						60
#define ADK_DB_CODE_OF_COUNTRY_CODE								61
#define ADK_DB_CUSTOMERPAYMENT									62
#define ADK_DB_CODE_OF_ADJUSTMENT_CODE							63
#define ADK_DB_SUPPLIERPAYMENT									64
#define ADK_DB_VERIFICATION_HEAD								65
#define ADK_DB_VERIFICATION_ROW									66
#define ADK_DB_CODE_OF_BOOKINGYEAR								67
#define ADK_DB_CODE_OF_DISCOUNT_CODE							68
#define ADK_DB_CONTACT											69
#define ADK_DB_CODE_OF_CONTACT_TITLES							70
#define ADK_DB_CODE_OF_CONTACT_GROUPS							71
#define ADK_DB_CODE_OF_CONTACT_GROUP_CONTACTS					72
#define ADK_DB_TAX_REDUCTION									73
#define ADK_DB_AGREEMENT_HEAD									74
#define ADK_DB_AGREEMENT_ROW									75
#define ADK_DB_TAX_REDUCTION_ORDER								76
#define ADK_DB_TAX_REDUCTION_AVTAL								77
#define ADK_DB_VERIFICATION_SERIES								78
#define ADK_DB_BOOKKEEPINGHIST									79
#define ADK_DB_PERIODIC_ADJUSTMENT								80
#define ADK_DB_CUSTOMER_ARTICLE									81
#define ADK_DB_ATTACHMENT_INFO									82
#define ADK_DB_TAX_REDUCTION_TYPES								83
#define ADK_DB_PRM3												84
```

## License

[![License](http://img.shields.io/:license-mit-blue.svg?style=flat-square)](http://badges.mit-license.org)

- **[MIT license](http://opensource.org/licenses/mit-license.php)**
