"""
Copyright (c) 2023 Sean Yeatts. All rights reserved.
"""

# ----------------------
# I. Module Dependencies
# ----------------------

# 3rd party imports
from quickpathstr       import Filepath     # syntax standardization helper

# Standard imports
import os                                   # manipulate folders
import win32com.client  as win              # COM object handling
import pythoncom        as pycom            # used in conjunction with win32com.client
import subprocess       as subproc          # quick process disconnect

# Project imports
from .utilities        import singleton     # singleton enforcement
from .utilities        import profile       # process timing


# ---------------------
# II. Import Definition
# ---------------------

__all__ = [
    "solidworks",
    "vault",
    "SWDocument",
    "Filepath",
    "profile"
]


# ------------
# III. Classes
# ------------

# Helper class
class SWDocument:
    """
    Document
    --------
    Convenience wrapper for SolidWorks IModelDoc2 objects.
    """
    # Special Methods
    # ---------------
    def __init__(self, swobj):
        self.filepath   = Filepath(swobj.GetPathName)   # Filepath
        self.swobj      = swobj                         # IModelDoc2
        

# Core class
@singleton
class SolidWorks:
    """
    SolidWorks
    ---------
    Wrapper for SolidWorks API. Represents the top-level application.
    """
    # Attributes
    # ----------
    client  = None
    version = None

    # Pubic Methods
    # -------------
    @profile
    def connect(self, version: int=2021, visible: bool=False):
        '''Creates a connection to the SolidWorks process.'''
        print(f"Connecting to SolidWorks client...")

        # Instantiate SolidWorks application via win32com dispatch (w/o concrete CLSID)
        self.version = version
        if not self.client:                                                                     # if a client is not defined...
            self.client = win.Dispatch("SldWorks.Application.%d" % (int(self.version)-1992))    # connect to new client  
            self.client.Visible = visible                                                       # make application visible
        else:                                                                                   # ...else terminal warning
            print(f"SolidWorks client connection is already established!")

    @profile
    def disconnect(self):
        '''Terminates connection to the SolidWorks process.'''
        print(f"Terminating SolidWorks process...")
        subproc.call(f"Taskkill /IM SLDWORKS.exe /F")
        # No follow-up terminal message necessary; subproc.call() auto-responds

    def open(self, filepath: Filepath) -> SWDocument:
        '''Opens a document using the specified complete path.'''
        print(f"Opening: {filepath.name}")

        # Define COM VARIANT arguments
        doc_type = None
        match filepath.extension:
            case '.SLDPRT':
                doc_type = win.VARIANT(pycom.VT_I4, 1)
            case '.SLDASM':
                doc_type = win.VARIANT(pycom.VT_I4, 2)
            case '.SLDDRW':
                doc_type = win.VARIANT(pycom.VT_I4, 3)

        file        = win.VARIANT(pycom.VT_BSTR,    filepath.complete)
        options     = win.VARIANT(pycom.VT_I4,      1)
        config      = win.VARIANT(pycom.VT_BSTR,    None)
        errors      = win.VARIANT(pycom.VT_BYREF |  pycom.VT_I4, 2)
        warnings    = win.VARIANT(pycom.VT_BYREF |  pycom.VT_I4, 128)

        # Execute SW-API method
        swobj = self.client.OpenDoc6(file, doc_type, options, config, errors, warnings)

        # Return Document
        return SWDocument(swobj)
    
    def safeclose(self, document: SWDocument):
        '''Closes the target document ( WITH rebuild and save methods ).'''
        self.rebuild(document)
        self.save(document)

        print(f"Closing: {document.filepath.name}")

        # Execute SW-API method
        self.client.CloseDoc(document.filepath.complete)

    def close(self, document: SWDocument):
        '''Closes the target document ( WITHOUT rebuild and save methods ).'''
        print(f"Closing: {document.filepath.name}")

        # Execute SW-API method
        self.client.CloseDoc(document.filepath.complete)

    def save(self, document: SWDocument):
        '''Saves the target document.'''
        print(f"Saving: {document.filepath.name}")

        # Define COM VARIANT arguments
        options     = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)
        errors      = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)
        warnings    = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)

        # Execute SW-API method
        document.swobj.Save3(options, errors, warnings)

    def rebuild(self, document: SWDocument):
        '''Rebuilds the target document.'''
        print(f"Rebuilding: {document.filepath.name}")
        
        # Define COM VARIANT arguments
        arg1 = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, False)

        # Execute SW-API method
        document.swobj.ForceRebuild3(arg1)

    def export(self, document: SWDocument, as_type: str="PNG"):
        '''Exports the target document as the prescribed file type.'''

        # Guard against incompatible types
        if document.filepath.extension == '.SLDDRW' and as_type != 'pdf':
            print(f"Drawings can only be exported as PDFs")
            return -1

        # Format components
        extension   = str("." + as_type)
        desktop     = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        destination = str(desktop + fr"\SolidWrap Exports")
        file        = Filepath(f"{destination}\{document.filepath.root}{extension}")

        # Make export directory
        if not os.path.exists(destination):
            os.makedirs(destination)

        print(f"Exporting: {file.name}")

        # Ignore staging for drawings
        if document.filepath.extension != '.SLDDRW':
            self.stage(document)

        # Define COM VARIANT arguments
        arg1 = win.VARIANT(pycom.VT_DISPATCH, None)
        arg2 = win.VARIANT(pycom.VT_BOOL, 0) 
        arg3 = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 0)
        arg4 = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 0)

        # Execute SW-API method
        document.swobj.Extension.SaveAs2(file.complete, 0, 1, arg1, "", arg2, arg3, arg4)

    def freeze(self, document: SWDocument):
        '''Freezes the target document's feature tree.'''
        print(f"Freezing: {document.filepath.name}")

        # Define COM VARIANT arguments
        setting = win.VARIANT(pycom.VT_I4, 461)

        # Execute SW-API method - show freeze bar
        self.client.SetUserPreferenceToggle(setting, True)

        # Get last feature in Feature Tree
        feature = get_last_feature(document)

        # Define COM VARIANT arguments
        position = win.VARIANT(pycom.VT_I4, 3)

        # Execute SW-API method - move freeze bar to end of feature tree
        document.swobj.FeatureManager.EditFreeze(position, feature.Name, True)

    def declutter(self, document: SWDocument, declutter: bool=True):
        '''Hides/Shows all of the target document's sketches, reference geometry, etc.'''
        print(f"Decluttering: {document.filepath.name}")

        # Define COM VARIANT arguments
        setting = win.VARIANT(pycom.VT_I4, 198) # swUserPreferenceToggle_e.swViewDisplayHideAllTypes
        
        # Execute SW-API Method: SetUserPreferenceToggle - hide all types (planes, sketches, etc.)
        document.swobj.Extension.SetUserPreferenceToggle(setting, 0, declutter)
        
    def stage(self, document: SWDocument):
        '''Declutters viewport and orients an isometric document view.'''
        print(f"Staging: {document.filepath.name}")
        self.declutter(document=document)

        # Execute SW-API Method: ShowNamedView2 - orient document to isometric view
        document.swobj.ShowNamedView2("Isometric", 7)

        # Execute SW-API Method: ViewZoomtofit2 - center document in viewport
        document.swobj.ViewZoomtofit2()

        # Execute SW-API Method: InsertScene - force background to plain white
        document.swobj.Extension.InsertScene(fr"\scenes\01 basic scenes\11 white kitchen.p2s")


# Core class
@singleton
class Vault:
    """
    Vault
    ---------
    Wrapper for SolidWorks PDM API. Represents PDM Vault.
    """
    # Attributes
    # ----------
    client      = None  # win32com application
    name        = None  # PDM Vault name (ex. "My_Vault")
    auth_state  = False # login credential authorization flag

    # Public Methods
    # --------------
    @profile
    def connect(self, name: str="My_Vault"):
        '''Creates a connection to the PDM Vault.'''
        print(f"Connecting to PDM...")

        # Instantiate PDM Vault via win32com dispatch (w/o concrete CLSID)
        self.name = name
        if not self.client:                                     # if a client is not defined...
            self.client = win.Dispatch("ConisioLib.EdmVault")   # connect to client
        else:                                                   # ...else terminal warning
            print(f"PDM connection is already established!")
        self.authenticate()

    @profile
    def authenticate(self):
        '''Authenticates login credentials for PDM Vault.'''
        print(f"Authentiating PDM credentials...")

        # Attempt login & flag authentication state
        if not self.client.IsLoggedIn:
            self.client.LoginAuto(self.name, 0)
            self.auth_state = True

    def checkout(self, filepath: Filepath):
        '''Checks out a file from the PDM Vault.'''
        print(f"Check Out: {filepath.name}")

        # Get PDM-API objects
        directory = self.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                        # IEdmFile

        # Execute PDM-API method
        if not file.IsLocked:                       # if file is not checked out...
            file.LockFile(directory.ID, 0)          # check out file
        else:
            print(f"File is already checked out!")  # ...else terminal warning
    
    # WIP
    def batch_checkout(self, files):
        '''Checks out a collection of files from the PDM Vault.'''
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
        '''Checks in a file to the PDM Vault.'''
        print(f"Check In: {filepath.name}")

        # Get PDM-API objects
        directory = vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder5
        file = directory.GetFile(filepath.name)                         # IEdmFile5

        # Execute PDM-API method
        if file.IsLocked:                           # if file is not checked in...
            file.UnlockFile(0, message)             # check in file
        else:
            print(f"File is already checked in!")   # ...else terminal warning

    # WIP
    def batch_checkin(self, files, message: str="SolidWrap Automated Check In"):
        '''Checks in a collection of files to the PDM Vault.'''
        for file in files:
            self.checkin(filepath=file)

    # WIP
    def change_state(self, filepath: Filepath, state: str="WIP", message: str="SolidWrap Automated State Change"):
        '''Changes a file's PDM state to the prescribed value, if allowed.'''
        print(f"Change State: {filepath.name}")

        # directory   = self.client.GetFolderFromPath(filepath.directory) # IEdmFolder5
        # folder_id   = directory.ID                                      # IEdmFolder5.ID (int?)
        # file        = directory.GetFile(filepath.name)                  # IEdmFile
        # file.ChangeState(state, folder_id, message, 0, 0)

    def get_revision(self, filepath: Filepath) -> str:
        '''Returns current revision of the target file.'''
        
        # Get PDM-API objects
        directory = vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder5
        file = directory.GetFile(filepath.name)                         # IEdmFile5
        return file.CurrentRevision
    
    def is_checked_out(self, filepath: Filepath) -> bool:
        '''Returns checkout state of the target file.'''
        
        # Get PDM-API objects
        directory = vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder5
        file = directory.GetFile(filepath.name)                         # IEdmFile5
        
        if file.IsLocked:
            return True
        else:
            return False

    def get_size(self, filepath: Filepath) -> int:
        '''Returns size of the target file (MB).'''

        # Get PDM-API objects
        directory = self.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                        # IEdmFile
        return int(file.GetLocalFileSize2(directory.ID)/1000)

    # WIP
    def get_pdm_state(self, filepath: Filepath) -> str:
        '''Returns PDM state of the target file.'''
        
        # Get PDM-API objects
        directory = vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder5
        file = directory.GetFile(filepath.name)                         # IEdmFile5
        return file.CurrentState.GetFirstTransitionPosition()


# --------------
# IV. Singletons
# --------------

vault = Vault()
solidworks = SolidWorks()


# ------------
# V. Functions
# ------------

def get_last_feature(document: SWDocument):
    '''Gets the last feature in the document's Feature Tree.'''

    # Get Feature Manager
    manager = document.swobj.FeatureManager

    # Pre-iteration setup
    count = manager.GetFeatureCount(True)
    tree = manager.GetFeatureTreeRootItem2(0)

    # Iterate through tree to get last item (yes, it has to be done this way)
    feature = tree.GetFirstChild
    for item in range(count):           # for each feature...
        if feature.GetNext:             # if valid item...
            feature = feature.GetNext   # ...then get item
    return feature.Object               # return the feature
