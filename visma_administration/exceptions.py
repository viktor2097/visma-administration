class VismaAPIError(Exception):
    """Base class for exceptions in this module."""

    pass


class InvalidFieldError(VismaAPIError):
    """Exception raised for errors in the input when invalid field is used."""

    pass


class ConnectionError(VismaAPIError):
    """Exception raised for errors while connecting to the database."""

    pass


class CompanyNotFoundError(VismaAPIError):
    """Exception raised when company is not found."""

    pass


class CredentialError(VismaAPIError):
    """Exception raised for errors in retrieving credentials."""

    pass
