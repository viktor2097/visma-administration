import logging
import os
from collections import namedtuple
from functools import partial
from enum import Enum
from os import sys
from winreg import HKEY_LOCAL_MACHINE, OpenKey, QueryValueEx

import clr
from System import Boolean, Double, String, DateTime

# Load the C# DLL
with OpenKey(
    HKEY_LOCAL_MACHINE,
    "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\SpcsAdm.Exe",
) as key:
    adk = QueryValueEx(key, "AdkDll")[0]
    clr.AddReference(adk + "\\AdkNet4Wrapper.dll")

from AdkNet4Wrapper import Api


class InvalidFilter(Exception):
    pass


class VismaAPI:
    """
    Call VismaAPI to open connection with a database (visma company folder)

    Example:
        VismaAPI(common_path="Z:\\Gemensamma filer", company_path="Z:\\Företag\\FTG9")

    The API may be accessed directly with the .api attribute,
    this property exposes the C# DLL and gives you access to all functionality
    defined in AdkNet4Wrapper.dll and Adk.h

    Example:
        VismaAPI().api.AdkSetFilter() # AdkSetFilter is a method of the dll.
    """

    _API = None

    def __init__(self, *args, **kwargs):
        self.adk = str()
        self.common_path = str()
        self.company_path = str()
        self.load_registry_keys()

        if "common_path" in kwargs or "company_path" in kwargs:
            self.common_path = kwargs.pop("common_path", self.common_path)
            self.company_path = kwargs.pop("company_path", self.company_path)

        # credentials for AdkOpen
        login = self.get_login_credentials()
        self.username = login.username
        self.password = login.password

        self.__class__._API = self.api

    def __del__(self):
        self.api.AdkClose()

    def load_registry_keys(self):
        with OpenKey(
            HKEY_LOCAL_MACHINE,
            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\SpcsAdm.Exe",
        ) as key:
            self.common_path = QueryValueEx(key, "CommonFiles")[0]
            self.company_path = QueryValueEx(key, "DefaultCompanyPath")[0]

    @property
    def api(self):
        if self._API:
            return self._API

        error = Api.AdkOpen2(
            self.common_path, self.company_path, self.username, self.password
        )
        if error.lRc != Api.ADKE_OK:
            error_message = Api.AdkGetErrorText(error, Api.ADK_ERROR_TEXT_TYPE.elRc)
            logging.error(error_message)
            sys.exit(1)

        return Api

    @staticmethod
    def get_login_credentials() -> namedtuple:
        Credentials = namedtuple("Credentials", ["username", "password"])

        try:
            username = os.environ["visma_username"]
            password = os.environ["visma_password"]
        except KeyError:
            logging.error("Set visma_username & visma_password environment variables")
            sys.exit(1)
        return Credentials(username=username, password=password)

    @staticmethod
    def db_fields():
        """
        Returns db fields defined in the DLL and Adk.h
        """
        fields = [field for field in Api.__dict__ if field.startswith("ADK_DB")]
        return fields

    @staticmethod
    def fields_for_db_field(db_field):
        """
        Example valid db_field is ADK_DB_SUPPLIER ( last word being the keyword )

        Fields of type ADK_DB_SUPPLIER are defined as ADK_SUPPLIER_FIELDNAME

        Function returns a list of names of associated fields of type db_field
        """

        if db_field not in VismaAPI.db_fields():
            logging.error(f"{db_field} is not listed as a db field in Adk.h")

        fields = [
            field
            for field in Api.__dict__
            if field.startswith("ADK_" + VismaAPI.field_without_db_prefix(db_field))
        ]
        return fields

    @staticmethod
    def field_without_db_prefix(db_field):
        return db_field.replace("ADK_DB_", "")


class _DBField:

    DB_NAME = None

    def __init__(self):
        self.api = VismaAPI._API
        self.pdata = _Pdata(
            self.__class__.DB_NAME,
            self.api.AdkCreateData(getattr(self.api, self.DB_NAME)),
        )

    def set_filter(self, **kwargs):
        """
        Apply filter to self.pdata based on filter provided to kwargs.
        Currently only supports filtering on one field and picks the last provided.

        Example:

            Supplier().filter(A="a", B="b") # Only applies filtering on B
            # B must be a valid field on ADK_DB_SUPPLIER
            
        """
        for field, filter_term in kwargs.items():
            if field not in VismaAPI.fields_for_db_field(self.__class__.DB_NAME):
                raise Exception("Field not found.")

            error = self.api.AdkSetFilter(
                self.pdata.data, getattr(self.api, field), filter_term, 0
            )
            if error.lRc != self.api.ADKE_OK:
                raise InvalidFilter

    def get(self, **kwargs):
        """
        Returns a single object, or returns an exception.
        """
        self.set_filter(**kwargs)
        error = self.api.AdkFirstEx(self.pdata.data, False)
        if error.lRc != self.api.ADKE_OK:
            error_message = Api.AdkGetErrorText(
                error, self.api.ADK_ERROR_TEXT_TYPE.elRc
            )
            raise Exception(error_message)

        return self.pdata

    def filter(self, **kwargs):
        """
        Returns multiple objects with a generator
        """
        try:
            self.get(**kwargs)
            yield self.pdata
        except Exception:
            return

        # while (
        #    error := self.api.AdkNextEx(self.pdata.data, False).lRc == self.api.ADKE_OK
        # ) :

        while True:
            error = self.api.AdkNextEx(self.pdata.data, False).lRc
            if error != self.api.ADKE_OK:
                break

            yield self.pdata


class _Pdata(object):
    """
    Wrapper for pdata objects
    Exposes fields of the db_name type

    Access and set fields like normal instance attributes

    Example:
    
        # hello is an instance of Pdata which data is of type ADK_DB_SUPPLIER
        hello = Supplier().get(ADK_SUPPLIER_NAME="hello")

        # Access a field on hello
        hello.adk_supplier_name

        # Set a field on hello
        hello.adk_supplier_name = "hello1"
        hello.save()

        
    """

    def __init__(self, db_name, pdata):
        object.__setattr__(self, "api", VismaAPI._API)
        object.__setattr__(self, "db_name", db_name)
        object.__setattr__(self, "data", pdata)

    def __getattr__(self, key):
        if key.upper() not in VismaAPI.fields_for_db_field(self.db_name):
            raise AttributeError("Trying to access field that doesn't exist")

        type = self.get_type(key)
        default_arguments = (self.data, getattr(self.api, key.upper()))

        if type == self.api.ADK_FIELD_TYPE.eChar:
            return self.api.AdkGetStr(*default_arguments, String(""))[1]
        elif type == self.api.ADK_FIELD_TYPE.eDouble:
            return self.api.AdkGetDouble(*default_arguments, Double(0.0))[1]
        elif type == self.api.ADK_FIELD_TYPE.eBool:
            return self.api.AdkGetBool(*default_arguments, Boolean(0))[1]
        elif type == self.api.ADK_FIELD_TYPE.eDate:
            return self.api.AdkGetDate(*default_arguments, DateTime())[1]

    def __setattr__(self, key, value):
        if key.upper() not in VismaAPI.fields_for_db_field(self.db_name):
            raise AttributeError("Trying to access field that doesn't exist")

        # Make sure user assigns correct type to the field
        _type = self.get_type(key)
        if not self.assignment_types_are_equal(_type, value):
            raise Exception(f"Trying to assign incorrect type to {key}")

        default_arguments = (self.data, getattr(self.api, key.upper()))

        error = None
        if _type == self.api.ADK_FIELD_TYPE.eChar:
            error = self.api.AdkSetStr(*default_arguments, String(f"{value}"))[0]
        elif _type == self.api.ADK_FIELD_TYPE.eDouble:
            error = self.api.AdkSetDouble(*default_arguments, Double(value))
        elif _type == self.api.ADK_FIELD_TYPE.eBool:
            error = self.api.AdkSetBool(*default_arguments, Boolean(value))
        elif _type == self.api.ADK_FIELD_TYPE.eDate:
            error = self.api.AdkSetDate(*default_arguments, self.to_date(value))

        if error and error.lRc != self.api.ADKE_OK:
            error_message = Api.AdkGetErrorText(
                error, self.api.ADK_ERROR_TEXT_TYPE.elRc
            )
            raise Exception(error_message)

    def assignment_types_are_equal(self, field_type, input_type):
        """
        Check if assignment value is of same type as field
        For example:
            supplier.adk_supplier_name = "hello"

        adk_supplier_name is a string field and expects a string assignment
        """
        if field_type == self.api.ADK_FIELD_TYPE.eChar and isinstance(input_type, str):
            return True
        elif field_type == self.api.ADK_FIELD_TYPE.eDouble and isinstance(
            input_type, (float, int)
        ):
            return True
        elif field_type == self.api.ADK_FIELD_TYPE.eBool and isinstance(
            input_type, bool
        ):
            return True
        elif field_type == self.api.ADK_FIELD_TYPE.eDate and isinstance(
            input_type, datetime.datetime
        ):
            return True

        return False

    def to_date(self, date):
        """
        Turn datetime object into a C# datetime object
        """
        return DateTime(
            date.year, date.month, date.day, date.hour, date.minute, date.second
        )

    def get_type(self, key):
        type = self.api.AdkGetType(
            self.data, getattr(self.api, key.upper()), self.api.ADK_FIELD_TYPE.eUnused
        )
        return type[1]

    def save(self):
        self.api.AdkUpdate(self.data)

    def delete(self):
        self.api.AdkDeleteRecord(self.data)


# Look through fields defined on VismaAPI as ADK_DB_FIELDNAME
# Dynamically create classes as Fieldname
for db_field in VismaAPI.db_fields():
    class_name = VismaAPI.field_without_db_prefix(db_field).title()
    globals()[class_name] = type(class_name, (_DBField,), {"DB_NAME": db_field})
