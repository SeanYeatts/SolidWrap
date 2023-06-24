"""
Copyright (c) 2023 Sean Yeatts. All rights reserved.
"""
# ----------------------
# I. Module Dependencies
# ----------------------
import os                                   # manipulate windows folders
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
    def connect(self, version: int=2021, visible: bool=False):
        """
        Creates a connection to the SolidWorks process.
        """
        print(f"Connecting to SolidWorks client...")

        # Instantiate SolidWorks application via win32com dispatch (w/o concrete CLSID)
        self.version = version
        if not self.client:                                                                     # if a client is not defined...
            self.client = win.Dispatch("SldWorks.Application.%d" % (int(self.version)-1992))    # connect to new client  
            self.client.Visible = visible                                                       # make application visible
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

        # Execute SW-API method - activate document
        # arg1 = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 0)
        # self.client.ActivateDoc3(filepath.complete, False, 2, arg1)

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
        model.swobj.Save3(options, errors, warnings)

    def rebuild(self, model: Model):
        """
        Rebuilds the target model.
        """
        print(f"Rebuilding: {model.filepath.name}")
        
        # Define COM VARIANT arguments
        arg1 = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, False)

        # Execute SW-API method
        model.swobj.ForceRebuild3(arg1)

    def export(self, model: Model, as_type: str="PNG"):
        """
        Exports the target model as the prescribed file type.
        """
        # Format components
        extension   = str("." + as_type)
        desktop     = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        destination = str(desktop + fr"\Exports")
        file        = Filepath(f"{destination}\{model.filepath.root}{extension}")

        # Make export directory
        if not os.path.exists(destination):
            os.makedirs(destination)

        print(f"Exporting: {file.name}")

        self.stage(model)

        # Define COM VARIANT arguments
        arg1 = win.VARIANT(pycom.VT_DISPATCH, None)
        arg2 = win.VARIANT(pycom.VT_BOOL, 0) 
        arg3 = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 0)
        arg4 = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 0)

        # Execute SW-API method
        model.swobj.Extension.SaveAs2(file.complete, 0, 1, arg1, "", arg2, arg3, arg4)

    def freeze(self, model: Model):
        """
        Freezes the target model's feature tree.
        """
        print(f"Freezing: {model.filepath.name}")

        # Define COM VARIANT arguments
        setting = win.VARIANT(pycom.VT_I4, 461)

        # Execute SW-API method - show freeze bar
        self.client.SetUserPreferenceToggle(setting, True)

        # Get last feature in feature tree
        feature = model.swobj.Extension.GetLastFeatureAdded

        # Define COM VARIANT arguments
        position = win.VARIANT(pycom.VT_I4, 3)

        # Execute SW-API method - move freeze bar to end of feature tree
        model.swobj.FeatureManager.EditFreeze(position, feature.Name, True)

    def declutter(self, model: Model, declutter: bool=True):
        """
        Hides/Shows all of the target model's sketches, reference geometry, etc.
        """
        print(f"Decluttering: {model.filepath.name}")

        # Define COM VARIANT arguments
        setting = win.VARIANT(pycom.VT_I4, 198) # swUserPreferenceToggle_e.swViewDisplayHideAllTypes
        
        # Execute SW-API Method: SetUserPreferenceToggle - hide all types (planes, sketches, etc.)
        model.swobj.Extension.SetUserPreferenceToggle(setting, 0, declutter)
        
    def stage(self, model: Model):
        """
        Decultters viewport and orients an isometric model view.
        """
        print(f"Staging: {model.filepath.name}")
        self.declutter(model=model)

        # Execute SW-API Method: ShowNamedView2 - orient model to isometric view
        model.swobj.ShowNamedView2("Isometric", 7)

        # Execute SW-API Method: ViewZoomtofit2 - center model in viewport
        model.swobj.ViewZoomtofit2()

        # Execute SW-API Method: InsertScene - force background to plain white
        model.swobj.Extension.InsertScene(fr"\scenes\01 basic scenes\11 white kitchen.p2s")

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
        print(f"Connecting to PDM...")

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
        directory = self.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                         # IEdmFile

        # Execute PDM-API method
        if not file.IsLocked:                       # if file is not checked out...
            file.LockFile(directory.ID, 0)          # check out file
        else:
            print(f"File is already checked out!")  # ...else terminal warning
    
    def batch_checkout(self, files):
        """
        Checks out a collection of models from PDM Vault.
        """
        for file in files:
            self.checkout(filepath=file)
        
        # ---
        # WIP
        # ---
        # filepath = fr"C:\Goddard_Vault\Users\SYeatts\Scripts"
        # filename = fr"C:\Goddard_Vault\Users\SYeatts\ScriptsTest_Part_01.sldprt"

        # folder      = vault.client.GetFolderFromPath(filepath)  # IEdmFolder5
        # folder_id   = folder.ID                                 # IEdmFolder5.ID (int?)

        # for file in files:
        #     item = folder.GetFile(file.name)            # IEdmFile5
        #     file_id = item.ID                           # IEdmFile5.ID (int?)
        #     utility = vault.client.CreateUtility(12)    # (22) IEdmBatchChangeState (12) IEdmBatchGet
        #     # utility.AddFile(file_id, folder_id)
        #     utility.AddSelectionEx(vault.client, file_id, folder_id, item.CurrentVersion)
        
        # utility.CreateTree(0, 2)    # @param2 (2) Egcf_Lock
        # # utility.GetFiles()
        # print(f"{utility.FileCount}")

    def checkin(self, filepath: Filepath, message: str="SolidWrap Automated Check In"):
        """
        Checks in model to PDM Vault.
        """
        print(f"Check In: {filepath.name}")

        # Get PDM-API objects
        directory = self.client.GetFolderFromPath(filepath.directory)  # IEdmFolder5
        file = directory.GetFile(filepath.name)                         # IEdmFile5

        # Execute PDM-API method
        if file.IsLocked:                           # if file is not checked in...
            file.UnlockFile(0, message)             # check in file
        else:
            print(f"File is already checked in!")   # ...else terminal warning

    def batch_checkin(self, files, message: str="SolidWrap Automated Check In"):
        """
        Checks in a collection of models to PDM Vault.
        """
        for file in files:
            self.checkin(filepath=file)

    def change_state(self, filepath: Filepath, state: str="WIP", message: str="SolidWrap Automated State Change"):
        """
        Changes model's PDM state to prescribed value, if allowed.
        """
        # print(f"Change State: {filepath.name}")

        # directory   = self.client.GetFolderFromPath(filepath.directory) # IEdmFolder5
        # folder_id   = directory.ID                                      # IEdmFolder5.ID (int?)
        # file        = directory.GetFile(filepath.name)                  # IEdmFile
        # file.ChangeState(state, folder_id, message, 0, 0)

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
