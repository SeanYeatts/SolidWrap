'''Copyright (c) 2024 Sean Yeatts. All rights reserved.'''

from __future__ import annotations


# IMPORTS - STANDARD LIBRARY
import functools


# CONSTANTS
AUTOMATION_MESSAGE      = 'automated action using SolidWrap'
EXPORT_FOLDER_DEFAULT   = 'SolidWrap Exports'
SUBPROCESS_NAME         = f"Taskkill /IM SLDWORKS.exe /F"
VAULT_DISPATCH_KEY      = 'ConisioLib.EdmVault'


# FUNCTIONS
def compute_client_key(version: int) -> str:
    """
    Formats the appropriate SolidWorks application client name based on
    the specified release version.
    """
    calculated = int(version - 1992)    # this is how SW calculates its version
    return ("SldWorks.Application.%d" % (calculated))

# WIP
# def protected_call(func, *args, **kwargs):
#     """
#     Wraps a target function in a pre-defined try / except for predictable
#     handling of SolidWorks API call exceptions. Use for any function that
#     contains an API call.
#     """
#     try:
#         return func(*args, **kwargs)
#     except Exception as exception:
#         print('ERROR: exception during SolidWorks API call')
#         print(exception)
#         print(traceback.print_exc())


# DECORATORS
def singleton(cls):
    """
    Gaurds an object against mutliple instantiation. New instances of the
    object will always return the singleton. This is DANGEROUS.
    """
    @functools.wraps(cls)                           # helper decorator to preserve class reference
    def wrapper(*args, **kwargs):
        if not wrapper.instance:                    # if an instance doesn't exist...
            wrapper.instance = cls(*args, **kwargs) # ...then store the instance
        return wrapper.instance
    wrapper.instance = None
    return wrapper                                  # return the instance


# CLASSES
class FileSize:

    # FUNDAMENTAL METHODS
    def __init__(self, size_bytes: float = 0.0) -> None:
        self.value = size_bytes
        self.suffix = self._calculate_suffix()
        self.concatenated = str(f"{self.value:.2f} {self.suffix}")

    # PRIVATE METHODS
    def _calculate_suffix(self) -> str:
        units = ["B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"]
        k = 1024.0
        if self.value < k:
            return "B"  # Return bytes if less than a kilobyte
        i = 0
        while self.value >= k and i < len(units) - 1:
            self.value /= k
            i += 1
        return f"{units[i]}"
