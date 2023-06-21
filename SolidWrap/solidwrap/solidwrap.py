"""
SolidWrap
=========

Copyright (c) 2023 Sean Yeatts. All rights reserved.

API Reference
-------------
"""

# Module Dependencies
import win32com.client  as win              # COM object handling
import pythoncom        as pycom            # used in conjunction with win32com.client
import subprocess       as subproc          # quick process disconnect
from quickpathstr       import Filepath     # file nomenclature standardization

# SW Type Libraries
import  rsc.swtypes as solidtype                    # type hinting
from    rsc.swconst import constants as solidconst  # explicit const names for readability

class ModelDoc:
    """
    ModelDoc
    --------
    Convenience wrapper for SolidWorks IModelDoc2 objects.
    """
    # Special Methods
    # ---------------
    def __init__(self, model: solidtype.IModelDoc2):
        self.model = model
        self.filepath = Filepath(model.GetPathName)

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
    vault   = None
    debug   = None

    # Special Methods
    # ---------------
    # Singleton protection
    # def __new__(cls):
    #     try:
    #         cls.__instance is None
    #     except:
    #         print(f"{cls.__instance.__str__} instance already exists!")
    #     else:
    #         cls.__instance = object.__new__(cls)
    #     finally:
    #         return cls.__instance

    # Initialize class Attribtues and connect to SW process
    def __init__(self, version: int, pdm_vault: str, debug=False) -> None:
        self.version = version
        self.debug = debug
        self.vault_connect(pdm_vault)
        self.connect()

    # Automatically terminate SW
    def __del__(self):
        if not self.debug:
            self.disconnect()

    # Public Methods
    # --------------
    # Connect to SW process (via win32com dispatch)
    def connect(self):
        """Creates a connection to the SolidWorks application."""
        # Instantiate SolidWorks via COM using concrete CLSID
        print(f"Establishing client connection...")
        if not self.client:
            self.client: solidtype.ISldWorks = win.Dispatch(solidtype.ISldWorks.coclass_clsid)
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

    def open_model(self, filepath: Filepath) -> ModelDoc:
        """Opens a model using the specified complete path."""
        print(f"Opening: {filepath.name}")

        # Define COM VARIANT arguments
        file        = win.VARIANT(pycom.VT_BSTR,    filepath.complete)
        doc_type    = win.VARIANT(pycom.VT_I4,      solidconst.swDocPART)
        options     = win.VARIANT(pycom.VT_I4,      solidconst.swOpenDocOptions_Silent)
        config      = win.VARIANT(pycom.VT_BSTR,    None)
        errors      = win.VARIANT(pycom.VT_BYREF |  pycom.VT_I4, solidconst.swFileNotFoundError)
        warnings    = win.VARIANT(pycom.VT_BYREF |  pycom.VT_I4, solidconst.swFileLoadWarning_AlreadyOpen)

        # Execute SW-API function call
        raw_model = self.client.OpenDoc6(file, doc_type, options, config, errors, warnings)
        
        # Return ModelDoc (encapsulated IModelDoc2)
        return ModelDoc(raw_model)

    def close(self, model: ModelDoc):
        """Closes the target model ( WITH rebuild() and save() calls )."""
        # Execute preparatory steps
        self.rebuild(model)
        self.save(model)
        
        print(f"Closing: {model.filepath.name}")
        # Execute SW-API function call
        self.client.CloseDoc(model.filepath.complete)

    def force_close(self, model: ModelDoc):
        """Closes the target model ( WITHOUT rebuild() and save() calls )."""
        # Execute SW-API function call
        self.client.CloseDoc(model.filepath.complete)

    def save(self, model: ModelDoc):
        """Saves the target model."""
        print(f"Saving: {model.filepath.name}")

        # Define COM VARIANT arguments
        options     = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, solidconst.swSaveAsOptions_Silent)
        errors      = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, solidconst.swGenericSaveError)
        warnings    = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, solidconst.swFileSaveWarning_RebuildError)

        # Execute SW-API function call
        model.model.Save3(options, errors, warnings)

    def rebuild(self, model: ModelDoc):
        """Rebuilds the target model."""
        print(f"Rebuilding: {model.filepath.name}")
        
        # Define COM VARIANT arguments
        arg1 = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, False)

        # Execute SW-API function call
        model.model.ForceRebuild3(arg1)

    # PDM Methods
    # --------------
    def vault_connect(self, vault_name):
        """Creates a connection to the PDM vault."""
        # Instantiate Vaule via COM using process name
        if not self.vault:
            self.vault = win.Dispatch("ConisioLib.EdmVault")
            print(f"Vault connection established!")
        else:
            print(f"Connection already established!")
        
        print(f"Authenticating PDM credentials...")
        self.vault.LoginAuto(vault_name, 0)

        if self.vault.IsLoggedIn:
            print(f"Login successful!")
        else:
            print(f"Login attempt failed!")
    
    def checkout(self, filepath: Filepath):

        parent_folder = self.vault.GetFolderFromPath(filepath.directory)    # IEdmFolder5
        file = parent_folder.GetFile(filepath.name)                         # IEdmFile5

        print(f"Checking out file: {filepath.name}")
        if not file.IsLocked:
            parentID = parent_folder.ID
            parentWND = 0
            file.LockFile(parentID, parentWND)
            print(f"File checked out!")
        else:
            print(f"File already checked out!")
    
    def checkin(self, filepath: Filepath):

        parent_folder = self.vault.GetFolderFromPath(filepath.directory)    # IEdmFolder5
        file = parent_folder.GetFile(filepath.name)                         # IEdmFile5

        print(f"Checking out file: {filepath.name}")
        if file.IsLocked:
            parentID = parent_folder.ID
            parentWND = 0
            file.UnlockFile(parentID, parentWND)
            print(f"File checked out!")
        else:
            print(f"File already checked out!")
