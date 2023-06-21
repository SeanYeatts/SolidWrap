
# ----------------------
# I. Module Dependencies
# ----------------------
import win32com.client  as win              # COM object handling
import pythoncom        as pycom            # used in conjunction with win32com.client
import subprocess       as subproc          # quick process disconnect
from quickpathstr       import Filepath     # file nomenclature standardization


# ------------------------
# II. Forward Declarations - circumvents scope annoyances
# ------------------------
class Model:        pass
class Vault:        pass
class SolidWrap:    pass


# ------------
# III. Classes
# ------------
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

    # Pubic Methods
    # -------------
    def connect(self, version: int=2021):
        """
        Creates a connection to the SolidWorks process.
        """
        print(f"Establishing connection to SolidWorks client...")

        # Instantiate SolidWorks application via win32com dispatch (w/o concrete CLSID)
        self.version = version
        if not self.client:                                                                     # if a client is not defined...
            self.client = win.Dispatch("SldWorks.Application.%d" % (int(self.version)-1992))    # connect to client
            self.client.Visible = True                                                          # make application visible
        else:                                                                                   # ...else terminal warning
            print(f"SolidWorks client connection is already established!")

    def disconnect(self):
        """
        Terminates connection to the SolidWorks process.
        """
        print(f"Terminating SolidWorks process...")
        subproc.call(f"Taskkill /IM SLDWORKS.exe /F")
        # No follow-up terminal message necessary; subproc.call() auto-responds

    def open(self, filepath: Filepath) -> Model:
        """
        Opens a model using the specified complete path.
        """
        print(f"Opening: {filepath.name}")

        # Define COM VARIANT arguments
        file        = win.VARIANT(pycom.VT_BSTR,    filepath.complete)
        doc_type    = win.VARIANT(pycom.VT_I4,      1)
        options     = win.VARIANT(pycom.VT_I4,      1)
        config      = win.VARIANT(pycom.VT_BSTR,    None)
        errors      = win.VARIANT(pycom.VT_BYREF |  pycom.VT_I4, 2)
        warnings    = win.VARIANT(pycom.VT_BYREF |  pycom.VT_I4, 128)

        # Execute SW-API method
        raw_model = self.client.OpenDoc6(file, doc_type, options, config, errors, warnings)

        # Return Model
        return Model(raw_model)
    
    def safeclose(self, model: Model):
        """
        Closes the target model ( WITH rebuild and save methods ).
        """
        self.rebuild(model)
        self.save(model)

        print(f"Closing: {model.filepath.name}")

        # Execute SW-API method
        self.client.CloseDoc(model.filepath.complete)

    def close(self, model: Model):
        """
        Closes the target model ( WITHOUT rebuild and save methods ).
        """
        print(f"Closing: {model.filepath.name}")

        # Execute SW-API method
        self.client.CloseDoc(model.filepath.complete)

    def save(self, model: Model):
        """
        Saves the target model.
        """
        print(f"Saving: {model.filepath.name}")

        # Define COM VARIANT arguments
        options     = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)
        errors      = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)
        warnings    = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)

        # Execute SW-API method
        model.model.Save3(options, errors, warnings)

    def rebuild(self, model: Model):
        """
        Rebuilds the target model.
        """
        print(f"Rebuilding: {model.filepath.name}")
        
        # Define COM VARIANT arguments
        arg1 = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, False)

        # Execute SW-API method
        model.model.ForceRebuild3(arg1)


class Vault:
    """
    Vault
    ---------
    Wrapper for SolidWorks PDM API. Represents PDM Vault.
    """
    # Attributes
    # ----------
    client      = None  # win32com application
    name        = None  # PDM Vault name (ex. "Goddard_Vault")
    auth_state  = False # login credential authorization flag

    # Public Methods
    # --------------
    def connect(self, name: str="Goddard_Vault"):
        """
        Creates a connection to the PDM Vault.
        """
        print(f"Establishing connection to PDM...")

        # Instantiate PDM Vault via win32com dispatch (w/o concrete CLSID)
        self.name = name
        if not self.client:                                     # if a client is not defined...
            self.client = win.Dispatch("ConisioLib.EdmVault")   # connect to client
        else:                                                   # ...else terminal warning
            print(f"PDM connection is already established!")
        self.authenticate()

    def authenticate(self):
        """
        Authenticates login credentials for PDM Vault.
        """
        print(f"Authentiating PDM credentials...")

        # Attempt login & flag authentication state
        if not self.client.IsLoggedIn:
            self.client.LoginAuto(self.name, 0)
            self.auth_state = True

    def checkout(self, filepath: Filepath):
        """
        Checks out model from PDM Vault.
        """
        print(f"Check Out: {filepath.name}")

        # Get PDM-API objects
        directory = vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder5
        file = directory.GetFile(filepath.name)                         # IEdmFile5

        # Execute PDM-API method
        if not file.IsLocked:                       # if file is not checked out...
            file.LockFile(directory.ID, 0)          # check out file
        else:
            print(f"File is already checked out!")  # ...else terminal warning
    
    def checkin(self, filepath: Filepath, message: str="automated checkin"):
        """
        Checks in model to PDM Vault.
        """
        print(f"Check In: {filepath.name}")

        # Get PDM-API objects
        directory = vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder5
        file = directory.GetFile(filepath.name)                         # IEdmFile5

        # Execute PDM-API method
        if file.IsLocked:                           # if file is not checked in...
            file.UnlockFile(0, message)             # check in file
        else:
            print(f"File is already checked in!")   # ...else terminal warning

    def batch_checkout(self, filepath: Filepath):
        """
        Checks out a collection of models from PDM Vault.
        """

    def batch_checkin(self, filepath: Filepath, message: str="automated checkin"):
        """
        Checks in a collection of models to PDM Vault.
        """


class Model:
    """
    Model
    -----
    Convenience wrapper for SolidWorks IModelDoc2 objects.
    """
    # Special Methods
    # ---------------
    def __init__(self, model):
        self.filepath = Filepath(model.GetPathName) # Filepath
        self.swobj = model                          # IModelDoc2


# ---------------------
# IV. Global Singletons
# ---------------------
vault = Vault()
solidworks = SolidWrap()
