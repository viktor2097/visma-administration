import datetime
import logging
import os
import time
from collections import namedtuple
from contextlib import contextmanager
from decimal import Decimal
from os import sys
from winreg import HKEY_LOCAL_MACHINE, OpenKey, QueryValueEx

import clr
from System import Boolean, DateTime, Double, String, Int32

with OpenKey(
    HKEY_LOCAL_MACHINE,
    "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\SpcsAdm.Exe",
) as key:
    adk = QueryValueEx(key, "AdkDll")[0]
    clr.AddReference(adk + "\\AdkNet4Wrapper.dll")

from AdkNet4Wrapper import Api


class InvalidFilter(Exception):
    pass


class Visma:
    """
    Simple interface between Visma Administration and Python

    Example:
        # Add a company
        Visma.add_company(
            name="Business Inc",
            common_path="Y:\\Gemensamma filer",
            company_path="Y:\\Företag\\FTG10"
        )

        # This opens the company in Visma associated with Business Inc company name
        # you defined earlier
        with Visma.get_company_api("Business Inc") as api:
            supplier = api.supplier.get(adk_supplier_name="supplier name")
    """

    _active_company = None
    active_sessions = 0
    companies = {}

    def __init__(self, *args, **kwargs):
        super().__init__()
        self.company = kwargs["company"]

        self.available_fields = {
            self.field_without_db_prefix(field).lower(): field
            for field in self.db_fields()
        }

    @classmethod
    @contextmanager
    def get_company_api(cls, name):
        """
        Yields an object for provided company name,
        providing a simple api to read, update and delete records related to the company

        Call Visma.add_company() before using this function,
        If more than one company is configured, it waits for other
        requests on a specific company to finish before yielding the object
        """
        if name not in cls.companies:
            raise CompanyNotFoundError("Company not found. Consider adding it first.")

        instance = cls(company=cls.companies[name])
        company = cls.companies[name]["company_path"]
        try:
            if cls._active_company == company:
                cls.active_sessions += 1
                yield instance
            else:
                if cls.wait_for(cls.no_active_sessions):
                    cls.active_sessions += 1

                    # Close previous company
                    if cls._active_company:
                        Api.AdkClose()

                    instance.api  # calling this so it sets new active company
                    yield instance
                else:
                    cls.active_sessions += 1
                    raise TimeoutError("Took too long to obtain the company API.")
        finally:
            cls.active_sessions -= 1

    @staticmethod
    def wait_for(predicate, timeout=60):
        deadline = time.time() + timeout
        while time.time() < deadline:
            if predicate():
                return True
            time.sleep(0.1)
        return False

    @classmethod
    def no_active_sessions(cls):
        if cls.active_sessions == 0:
            return True
        return False

    @classmethod
    def add_company(cls, name, common_path, company_path, username=None, password=None):
        """
        Adds a company to VismaAPI.companies
        name is used with .get_company_api()
        """
        if not username or not password:
            login = cls.get_login_credentials()
            username = login.username
            password = login.password

        cls.companies[name] = {
            "common_path": common_path,
            "company_path": company_path,
            "username": username,
            "password": password,
        }

    def __getattr__(self, name):
        """
        Args:
            name: Name of DB_FIELD, E.g ADK_DB_SUPPLIER name is supplier ( ADK_DB_ is removed, and lowercase letters )

        Returns:
            A _DBField instance exposing methods for any ADK_DB_FIELD
        """
        if name in self.available_fields:
            return type(
                name.title(), (_DBField,), {"DB_NAME": self.available_fields[name]}
            )(api=self.api)
        else:
            raise InvalidFieldError(f"{name} is not a valid field.")

    @property
    def api(self):
        """
        Returns the Api object defined in DLL and opens a database connection if needed
        """
        if self.__class__._active_company == self.company["company_path"]:
            return Api
        else:
            self.__class__._active_company = self.company["company_path"]

        error = Api.AdkOpen2(
            self.company["common_path"],
            self.company["company_path"],
            self.company["username"],
            self.company["password"],
        )
        if error.lRc != Api.ADKE_OK:
            error_message = self._api.AdkGetErrorText(
                error, Api.ADK_ERROR_TEXT_TYPE.elRc
            )
            logging.error(error_message)
            raise ConnectionError(f"Error connecting to the Visma API: {error_message}")

        return Api

    @staticmethod
    def get_login_credentials() -> namedtuple:
        """
        Finds Visma credentials from visma_username and visma_password environment variables

        Returns:
            namedtuple containing username and password
        """
        Credentials = namedtuple("Credentials", ["username", "password"])

        try:
            username = os.environ["visma_username"]
            password = os.environ["visma_password"]
        except KeyError:
            logging.error(
                "provide username and password upon class instantiation,"
                "or set visma_username & visma_password environment variables"
            )
            raise CredentialError(
                "Failed to get Visma credentials from environment variables"
            )
        return Credentials(username=username, password=password)

    def db_fields(self):
        """
        Returns db fields defined in the DLL and Adk.h
        """
        fields = [field for field in self.api.__dict__ if field.startswith("ADK_DB")]
        return fields

    @staticmethod
    def field_without_db_prefix(db_field):
        return db_field.replace("ADK_DB_", "")


class _DBField:
    """
    Simple manager for pData.
    Used to filter and create new objects

    For instance,
        .filter(field_name="hello") creates a generator and yields a _Pdata object for each iteration
        .get() returns a single _Pdata object to work with
        .new() returns an empty _Pdata object
    """

    DB_NAME = None

    def __init__(self, api):
        self.api = api
        self.pdata = _Pdata(
            self.api,
            self.__class__.DB_NAME,
            self.api.AdkCreateData(getattr(self.api, self.DB_NAME)),
        )

    def set_filter(self, **kwargs):
        """
        Apply filter to self.pdata based on filter provided to kwargs.
        Currently only supports filtering on one field and picks the last provided.

        Example:

            supplier.filter(A="a", B="b") # Only applies filtering on B
            # B must be a valid field of ADK_DB_SUPPLIER

        """
        for field, filter_term in kwargs.items():
            field = field.upper()  # Fields in Visma are all uppercased
            try:
                self.api.AdkGetType(
                    self.pdata.data,
                    getattr(self.api, field),
                    self.api.ADK_FIELD_TYPE.eUnused,
                )
            except AttributeError:
                raise InvalidFieldError(
                    f"{field} is not a valid field of {self.__class__.DB_NAME}"
                )

            error = self.api.AdkSetFilter(
                self.pdata.data, getattr(self.api, field), filter_term, 0
            )
            if error.lRc != self.api.ADKE_OK:
                raise InvalidFilter

    def new(self):
        return self.pdata

    def get(self, include_rows=True, **kwargs):
        """
        Returns a single object, or raises an exception.
        """
        self.set_filter(**kwargs)
        error = self.api.AdkFirstEx(self.pdata.data, include_rows)
        if error.lRc != self.api.ADKE_OK:
            error_message = self.api.AdkGetErrorText(
                error, self.api.ADK_ERROR_TEXT_TYPE.elRc
            )
            raise Exception(error_message)

        return self.pdata

    def filter(self, include_rows=True, **kwargs):
        """
        Returns multiple objects with a generator
        """
        try:
            self.get(**kwargs)
            yield self.pdata
        except Exception:
            return

        while True:
            error = self.api.AdkNextEx(self.pdata.data, include_rows).lRc
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
        hello = visma.supplier.get(adk_supplier_name="hello")

        # Access a field on hello
        hello.adk_supplier_name

        # Set a field on hello
        hello.adk_supplier_name = "hello1"
        hello.save()
    """

    def __init__(
        self, api, db_name, pdata, parent_pdata=None, is_a_row=False, row_index=None
    ):
        object.__setattr__(self, "api", api)
        object.__setattr__(self, "db_name", db_name)
        object.__setattr__(self, "data", pdata)
        object.__setattr__(self, "is_a_row", is_a_row)
        object.__setattr__(self, "parent_pdata", parent_pdata)
        object.__setattr__(self, "row_index", row_index)

    def __del__(self):
        self.api.AdkDeleteStruct(self.data)

    def __getattr__(self, key):
        try:
            _type = self.get_type(key)
        except AttributeError:
            raise AttributeError(f"{key} is not a valid field of {self.db_name}")

        default_arguments = (self.data, getattr(self.api, key.upper()))

        if _type == self.api.ADK_FIELD_TYPE.eChar:
            return self.api.AdkGetStr(*default_arguments, String(""))[1]
        elif _type == self.api.ADK_FIELD_TYPE.eDouble:
            return self.api.AdkGetDouble(*default_arguments, Double(0.0))[1]
        elif _type == self.api.ADK_FIELD_TYPE.eBool:
            return self.api.AdkGetBool(*default_arguments, Boolean(0))[1]
        elif _type == self.api.ADK_FIELD_TYPE.eDate:
            return self.api.AdkGetDate(*default_arguments, DateTime())[1]

    def __setattr__(self, key, value):
        try:
            _type = self.get_type(key)
        except AttributeError:
            raise AttributeError(f"{key} is not a valid field of {self.db_name}")

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
            error_message = self.api.AdkGetErrorText(
                error, self.api.ADK_ERROR_TEXT_TYPE.elRc
            )
            raise Exception("field: {}, error: {}".format(key, error_message))

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
            input_type, (float, int, Decimal)
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
        error = self.api.AdkUpdate(self.data)
        if error.lRc != self.api.ADKE_OK:
            error_message = self.api.AdkGetErrorText(
                error, self.api.ADK_ERROR_TEXT_TYPE.elRc
            )
            raise Exception(error_message)

    def delete(self):
        if self.is_a_row:
            self.api.AdkDeleteRow(self.parent_pdata.data, self.row_index)
            return
        self.api.AdkDeleteRecord(self.data)

    def create(self):
        error = self.api.AdkAdd(self.data)
        if error.lRc != self.api.ADKE_OK:
            error_message = self.api.AdkGetErrorText(
                error, self.api.ADK_ERROR_TEXT_TYPE.elRc
            )
            raise Exception(error_message)

    def rows(self):
        """
        Returns a list of rows which are of type _Pdata
        You may access any fields, do assignments and delete rows like any other _Pdata objects
        """
        _row_db_id = self.api.AdkGetRowDataId(self.data, Int32(0))[1]
        _nrows_field_id = self.api.AdkGetNrowsFieldId(self.data, Int32(0))[1]
        _rows_field_id = self.api.AdkGetRowsFieldId(self.data, Int32(0))[1]
        nrows = self.api.AdkGetDouble(self.data, _nrows_field_id, Double(0.0))[1]

        _existing_rows = []
        for index in range(int(nrows)):
            data = self.api.AdkGetRowData(self.data, index, Int32(0))[1]
            _existing_rows.append(
                _Pdata(
                    self.api,
                    _row_db_id,
                    data,
                    parent_pdata=self,
                    is_a_row=True,
                    row_index=index + 1,
                )
            )

        return _existing_rows

    def create_rows(self, quantity=1):
        """
        Returns a list containg rows of _Pdata type.
        Example:
            rows = invoice.create_rows(2)
            rows[0].adk_ooi_... = "hello"
            rows[1].adk_ooi_... = "world"
            invoice.save()

        Args:
            quantity: number of rows to create
        """
        if quantity < 1:
            raise ValueError("New row quantity must be 1 or higher.")

        _row_db_id = self.api.AdkGetRowDataId(self.data, Int32(0))[1]
        _nrows_field_id = self.api.AdkGetNrowsFieldId(self.data, Int32(0))[1]
        _rows_field_id = self.api.AdkGetRowsFieldId(self.data, Int32(0))[1]
        nrows = self.api.AdkGetDouble(self.data, _nrows_field_id, Double(0.0))[1]

        _rows = self.api.AdkCreateDataRow(_row_db_id, int(nrows) + quantity)
        self.api.AdkSetDouble(self.data, _nrows_field_id, Double(nrows + quantity))
        self.api.AdkSetData(self.data, _rows_field_id, _rows)

        row_objects = []
        for index in range(quantity):
            actual_nrows_index = nrows + index
            row_objects.append(
                _Pdata(
                    self.api,
                    _row_db_id,
                    self.api.AdkGetDataRow(_rows, actual_nrows_index),
                    parent_pdata=self,
                    row_index=actual_nrows_index,
                )
            )

        return row_objects
