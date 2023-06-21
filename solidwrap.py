"""
SolidWrap
=========

A Python wrapper for the SolidWorks API (+ PDM).

Copyright (c) 2023 Sean Yeatts. All rights reserved.
"""

# Module Dependencies
# -------------------
import win32com.client  as win              # COM object handling
import pythoncom        as pycom            # used in conjunction with win32com.client
import subprocess       as subproc          # quick process disconnect
from quickpathstr       import Filepath     # file nomenclature standardization


# Classes
# -------
class Model:
    """
    Model
    -----
    Convenience wrapper for SolidWorks IModelDoc2 objects & their IEdmFile5 implementation.
    """
    # Special Methods
    # ---------------
    def __init__(self, model):
        self.filepath = Filepath(model.GetPathName) # Filepath
        self.model = model                          # IModelDoc2
        self.pdm_directory = vault.client.GetFolderFromPath(self.filepath.directory)    # IEdmFolder5
        self.entry = self.pdm_directory.GetFile(self.filepath.name)                     # IEdmFile5
        

class Vault:
    """
    PDM
    ---------
    Wrapper for PDM API.
    """
    # Attributes
    # ----------
    client          = None
    vault_name      = None
    b_authenticated = False

    # Public Methods
    # --------------
    def connect(self, vault_name: str):
        """Creates a connection to the PDM Vault."""
        print(f"Estabishing connection to PDM Vault...")
        if not self.client:
            # Instantiate core interface via COM using process ID
            self.client = win.Dispatch("ConisioLib.EdmVault")
            print(f"Connection successfully established!")
            self.authenticate(vault_name)
        else:
            print(f"Connection already established!")

    def authenticate(self, name: str):
        """Authenticates login for PDM Vault."""
        print(f"Authenticating PDM credentials...")
        if not self.client.IsLoggedIn:
            self.client.LoginAuto(name, 0)

        if self.client.IsLoggedIn:
            print(f"Login attempt successful!")
            self.vault_name = name
            self.b_authenticated = True
        else:
            print(f"Login attempt failed!")

    def checkout(self, filepath: Filepath):
        """Checks out model from PDM Vault."""

        directory = vault.client.GetFolderFromPath(filepath.directory) # IEdmFolder5
        file = directory.GetFile(filepath.name)                 # IEdmFile5

        print(f"Checking out file: {filepath.name}")
        if not file.IsLocked:
            parentID = directory.ID
            parentWND = 0
            file.LockFile(parentID, parentWND)
            print(f"Checkout successful!")
        else:
            print(f"File is already checked out!")

    def checkin(self, filepath: Filepath, message: str="SolidWrap automated workflow"):
        """Checks in model to PDM Vault."""

        directory = vault.client.GetFolderFromPath(filepath.directory) # IEdmFolder5
        file = directory.GetFile(filepath.name)                 # IEdmFile5

        print(f"Checking in file: {filepath.name}")
        if file.IsLocked:
            parentWND = 0
            comment = message
            file.UnlockFile(parentWND, comment)
            print(f"Checkin successful!")
        else:
            print(f"File is already checked in!")


class SolidWrap:
    """
    SolidWrap
    ---------
    Wrapper for SolidWorks API. Represents the top-level application.
    """
    # Attributes
    # ----------
    client  = None
    version = None
    debug   = None

    # Special Methods
    # ---------------
    # Automatically terminate SW
    # def __del__(self):
    #     if not self.debug:
    #         self.disconnect()

    # Public Methods
    # --------------
    # Connect to SW process (via win32com dispatch)
    def connect(self, version: int=2019, debug: bool=False):
        """Creates a connection to the SolidWorks application."""
        self.version = version
        self.debug = debug
        # Instantiate SolidWorks via COM using concrete CLSID
        print(f"Establishing SW client connection...")
        if not self.client:
            self.client = win.Dispatch("SldWorks.Application.%d" % (int(self.version)-1992))
            self.client.Visible = self.debug    # application visibility based on debug state
            print(f"Connection successfully established!")
        else:
            print(f"Connection already established!")

    # Disconnect from SW process
    def disconnect(self):
        """Terminates the SolidWorks process."""
        print(f"Terminating SolidWorks process...")
        subproc.call(f"Taskkill /IM SLDWORKS.exe /F")
        # No follow-up message necessary; subproc.call() auto-responds

    def open_model(self, filepath: Filepath) -> Model:
        """Opens a model using the specified complete path."""
        print(f"Opening: {filepath.name}")

        # Define COM VARIANT arguments
        file        = win.VARIANT(pycom.VT_BSTR,    filepath.complete)
        doc_type    = win.VARIANT(pycom.VT_I4,      1)
        options     = win.VARIANT(pycom.VT_I4,      1)
        config      = win.VARIANT(pycom.VT_BSTR,    None)
        errors      = win.VARIANT(pycom.VT_BYREF |  pycom.VT_I4, 2)
        warnings    = win.VARIANT(pycom.VT_BYREF |  pycom.VT_I4, 128)

        # Execute SW-API function call
        raw_model = self.client.OpenDoc6(file, doc_type, options, config, errors, warnings)
        
        # Return Model (encapsulated IModelDoc2)
        return Model(raw_model)

    def close(self, model: Model):
        """Closes the target model ( WITH rebuild() and save() calls )."""
        # Execute preparatory steps
        self.rebuild(model)
        self.save(model)
        
        print(f"Closing: {model.filepath.name}")
        # Execute SW-API function call
        self.client.CloseDoc(model.filepath.complete)

    def force_close(self, model: Model):
        """Closes the target model ( WITHOUT rebuild() and save() calls )."""
        # Execute SW-API function call
        self.client.CloseDoc(model.filepath.complete)

    def save(self, model: Model):
        """Saves the target model."""
        print(f"Saving: {model.filepath.name}")

        # Define COM VARIANT arguments
        options     = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)
        errors      = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)
        warnings    = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)

        # Execute SW-API function call
        model.model.Save3(options, errors, warnings)

    def rebuild(self, model: Model):
        """Rebuilds the target model."""
        print(f"Rebuilding: {model.filepath.name}")
        
        # Define COM VARIANT arguments
        arg1 = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, False)

        # Execute SW-API function call
        model.model.ForceRebuild3(arg1)

# Singleton Global Objects
app     = SolidWrap()   # SolidWorks process
vault   = Vault()       # PDM Vault
