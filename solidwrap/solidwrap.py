'''Copyright (c) 2024 Sean Yeatts. All rights reserved.'''

from __future__ import annotations


# IMPORTS - EXTERNAL
from quickpathstr import Filepath

# IMPORTS - STANDARD LIBRARY
import os                                   # file / folder manipulation
import pythoncom        as pycom            # used with win32com.client
import subprocess       as subproc          # quick process disconnect
import win32com.client  as win              # COM object handling
from pathlib import Path                    # file / folder manipulation
from typing import Any, Dict, List, Type    # static type checking

# IMPORTS - PROJECT
from containers import SWAuthState, SWDocType, SWExportFormat, SWDocument
from logger import log
from utilities import AUTOMATION_MESSAGE, EXPORT_FOLDER_DEFAULT
from utilities import SUBPROCESS_NAME, VAULT_DISPATCH_KEY
from utilities import compute_client_key, singleton


# SYMBOLS
__all__ = [
    'SolidWorks',
    'Vault',
    'SWDocument',
    'SWDocType',
    'SWExportFormat',
    'Filepath'
]


# CLASSES
@singleton
class SolidWorks:
    """
    SolidWorks
    ----------
    Wrapper for SolidWorks API. Represents the top-level client.
    """
    
    # CLASS ATTRIBUTES
    client: Any = None

    # FUNDAMENTAL METHODS
    def __init__(self, version: int = 2023) -> None:
        self.version = version

    # LIFETIME MANAGEMENT METHODS
    def connect(self, headless: bool = False) -> bool:
        '''Creates a connection to the SolidWorks client.'''
        log.critical(f"connecting to SolidWorks client ( {self.version} )")
        client_key = compute_client_key(self.version)
        
        # Attempt client dispatch
        log.info('establishing connection...')
        try:
            pycom.CoInitialize()
            if (com_object := win.Dispatch(client_key)):
                SolidWorks.client = com_object
                SolidWorks.client.visible = not headless
                log.info('connection successfully established')
                return True
        except Exception as exception:
            print('ERROR: failed to establish connection')
            print(f"ERROR: {exception}")
            return False

    def disconnect(self, silent: bool = True) -> bool:
        '''Terminates the SolidWorks connection.'''
        log.critical(f"disconnecting from SolidWorks client ( {self.version} )")
        
        # Attempt to terminate the subprocess
        try:
            if not silent:
                log.info('terminating subprocess')
                subproc.call(SUBPROCESS_NAME)
            pycom.CoUninitialize()
            SolidWorks.client = None
            log.info('connection successfully terminated')
            return True
        except Exception as exception:
            print('ERROR: failed to disconnect from SolidWorks client')
            print(f"ERROR: {exception}")
            return False

    # DOCUMENT MANAGEMENT METHODS
    def open(self, filepath: Filepath) -> SWDocument:
        log.info(f"opening document: '{filepath.name}'")

        # Evaluate document type by extension ( period removed )
        type_key: int = 0
        match str.upper(filepath.extension.replace('.', '')):
            case SWDocType.PART.value:
                type_key = 1
            case SWDocType.ASSEMBLY.value:
                type_key = 2
            case SWDocType.DRAWING.value:
                type_key = 3

        # Define COM VARIANT args
        source      = win.VARIANT(pycom.VT_BSTR,    filepath.complete)
        doc_type    = win.VARIANT(pycom.VT_I4,      type_key)
        options     = win.VARIANT(pycom.VT_I4,      1)
        config      = win.VARIANT(pycom.VT_BSTR,    None)
        errors      = win.VARIANT(pycom.VT_BYREF |  pycom.VT_I4, 2)
        warnings    = win.VARIANT(pycom.VT_BYREF |  pycom.VT_I4, 128)

        # SolidWorks API call
        swobj = SolidWorks.client.OpenDoc6(
            source, doc_type, options, config, errors, warnings
        )

        # Return wrapped com object
        return SWDocument(swobj)

    def close(self, document: SWDocument) -> None:
        log.info(f"closing document: '{document.source.name}'")

        # SolidWorks API call
        SolidWorks.client.CloseDoc(document.source.complete)

    def safeclose(self, document: SWDocument) -> None:
        self.rebuild(document, False)
        self.save(document)
        self.close(document)

    def save(self, document: SWDocument) -> None:
        log.info(f"saving document: '{document.source.name}'")

        # Define COM VARIANT args
        options     = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)
        errors      = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)
        warnings    = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)

        # SolidWorks API call
        document.swobj.Save3(options, errors, warnings)

    def rebuild(self, document: SWDocument, top_only: bool = False) -> None:
        log.info(f"rebuilding document: '{document.filepath.name}'")
        
        # Define COM VARIANT args
        top_only = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, top_only)

        # Execute SW-API method
        document.swobj.ForceRebuild3(top_only)

    # HIGH-LEVEL METHODS
    def export(self, document: SWDocument, as_type: SWExportFormat,
        destination: Filepath = None) -> None:
        
        # Technical setup
        output = prepare_export(document, as_type, destination)

        # Graphical setup
        self.stage(document)

        log.info(f"exporting document: '{output.name}'")

        # Define COM VARIANT args
        data        = win.VARIANT(pycom.VT_DISPATCH,   None)
        prefix      = win.VARIANT(pycom.VT_BOOL,       0) 
        errors      = win.VARIANT(pycom.VT_BYREF |     pycom.VT_I4, 0)
        warnings    = win.VARIANT(pycom.VT_BYREF |     pycom.VT_I4, 0)

        # Execute SW-API method
        document.swobj.Extension.SaveAs2(
            output.complete, 0, 1, data, "", prefix, errors, warnings
        )

    def stage(self, document: SWDocument) -> None:
        log.info(f"staging document: '{document.source.name}'")

        # Reject drawings
        if document.doctype == SWDocType.DRAWING:
            return None

        # Define COM VARIANT args & various other datapoints
        setting = win.VARIANT(pycom.VT_I4, 198)
        view_id = 7
        scene   = fr"\scenes\01 basic scenes\11 white kitchen.p2s"
        
        # Execute SW-API method - hide all types (planes, sketches, etc.)
        document.swobj.Extension.SetUserPreferenceToggle(setting, 0, True)

        # Execute SW-API method - orient document
        document.swobj.ShowNamedView2('Isometric', view_id)

        # Execute SW-API method - center document in viewport
        document.swobj.ViewZoomtofit2()

        # Execute SW-API method - force background to plain white
        document.swobj.Extension.InsertScene(scene)

    def freeze(self, document: SWDocument) -> None:
        log.info(f"freezing document: '{document.filepath.name}'")

        # Define COM VARIANT args
        setting     = win.VARIANT(pycom.VT_I4,  461)
        position    = win.VARIANT(pycom.VT_I4,  3)
        
        # Execute SW-API method - show freeze bar
        SolidWorks.client.SetUserPreferenceToggle(setting, True)

        # Get last feature in Feature Tree
        last_feature = self.get_last_feature(document)

        # Execute SW-API method - move freeze bar past last feature
        document.swobj.FeatureManager.EditFreeze(
            position, last_feature.Name, True
        )

    # LOW-LEVEL METHODS
    def get_last_feature(self, document: SWDocument) -> Any:
        log.debug('getting last feature in Feature Tree')
        
        # Gather necessary components of the Feature Tree
        manager = document.swobj.FeatureManager     # get Feature Manager
        tree = manager.GetFeatureTreeRoomItem2(0)   # get Feature Tree root
        count = manager.GetFeatureCount(True)       # get # of features

        # Iterate over tree to find last item ( yes, this is the only way... )
        feature = tree.GetFirstChild        # start with first feature
        for feature in range(count):        # iterate through features
            if feature.GetNext:             # if there's a next feature...
                feature = feature.GetNext   # ...re-assign current
        return feature.Object               # return the final feature


@singleton
class Vault:
    """
    VAULT
    -----
    Wrapper for SolidWorks API. Represents the PDM Vault.
    """

    # CLASS ATTRIBUTES
    client: Any = None

    # FUNDAMENTAL METHODS
    def __init__(self, name: str) -> None:
        self.name = name
        self.authorized = SWAuthState.UNAUTHORIZED
        
    # LIFETIME MANAGEMENT METHODS
    def connect(self) -> bool:
        '''Creates a connection to the PDM Vault client.'''
        log.critical(f"connecting to PDM Vault ( {self.name} )")
        
        # Attempt new client dispatch
        log.info('establishing connection...')
        try:
            pycom.CoInitialize()
            if (com_object := win.Dispatch(VAULT_DISPATCH_KEY)):
                Vault.client = com_object
                log.info('connected successfully established')
                return self.authorize()
        except Exception as exception:
            print('ERROR: failed to establish connection')
            print(f"ERROR: {exception}")
            return False

    def disconnect(self) -> bool:
        '''Unhooks from the PDM Vault process.'''
        log.critical(f"disconnecting from PDM Vault ( {self.name} )")

        # Attempt to disconnect from the process
        try:
            pycom.CoUninitialize()
            Vault.client = None
            self.authorized = SWAuthState.UNAUTHORIZED
            log.info('connection successfully terminated')
            return True
        except Exception as exception:
            print('ERROR: failed to disconnect from PDM Vault client')
            print(f"ERROR: {exception}")
            return False

    def authorize(self) -> bool:
        '''Authenticates login credentials for the PDM Vault.
        
        Returns:
            - ( bool ) : True if successful
        '''
        log.info('authenticating PDM credentials...')

        # Check for existing authorization
        if Vault.client.IsLoggedIn:
            log.info('credentials successfully authenticated')
            self.authorized = SWAuthState.AUTHORIZED
            return True
        
        # Attempt new authorization
        try:
            # Execute PDM-API method
            Vault.client.LoginAuto(self.name, 0)
            self.authorized = SWAuthState.AUTHORIZED
            log.info('credentials successfully authenticated')
            return True
        except Exception as exception:
            print('ERROR: failed authenticate login credentials')
            print(f"ERROR: {exception}")
            return False

    # FILE STATE METHODS
    def checkin(self, filepath: Filepath, comment: str = None) -> None:
        '''IMPORTANT: this method will not work if the file is currently
        open as a document in SolidWorks. Close the file before use.'''
        log.info(f"checking in file: '{filepath.name}'")

        # Get PDM-API objects
        directory = Vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                         # IEdmFile

        # Format a check in comment
        message = AUTOMATION_MESSAGE
        if comment:
            message = message + ': ' + comment

        # Check current document state         
        if not file.IsLocked:
            log.warning(f"file is already checked in: '{filepath.name}'")
            return None
        
        # Execute PDM-API method    
        file.UnlockFile(0, message)
            
    def checkout(self, filepath: Filepath) -> None:
        '''IMPORTANT: this method will not work if the file is currently
        open as a document in SolidWorks. Close the file before use.'''
        log.info(f"checking out file: '{filepath.name}'")

        # Get PDM-API objects
        directory = Vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                         # IEdmFile

        # Check current document state
        if file.IsLocked:
            log.warning(f"file is already checked out: '{filepath.name}'")
            return None
        
        # Execute PDM-API method
        file.LockFile(directory.ID, 0)

    def undo_checkout(self, filepath: Filepath) -> None:
        '''IMPORTANT: this method will not work if the file is currently
        open as a document in SolidWorks. Close the file before use.'''
        log.info(f"undoing check out: '{filepath.name}'")

        # Get PDM-API objects
        directory = Vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                         # IEdmFile

        # Check current document state
        if not file.IsLocked:
            log.warning(f"file is not checked out: '{filepath.name}'")
            return None
        
        # Execute PDM-API method
        file.UndoLockFile(0)

    # WORKFLOW STATE METHODS
    def change_state(self, filepath: Filepath, state: str,
        comment: str = None) -> None:
        log.info(f"changing workflow state: '{filepath.name}'")
        
        # GET FILE WORKFLOW DATA
        
        # log.debug(f"current state: '{asdf}'")
        # log.debug(f"new state: '{ghjk}'")

    # DATA RETRIEVAL METHODS
    def get_revision(self, filepath: Filepath) -> str:
        log.debug('release revision')

        # Get PDM-API objects
        directory = Vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder5
        file = directory.GetFile(filepath.name)                         # IEdmFile5
        return file.CurrentRevision


# FUNCTIONS
def prepare_export(document: SWDocument, target_format: SWExportFormat,
        destination: Filepath = None) -> Filepath:
    log.debug(f"preparing document for export: '{document.source.name}'")
    
    # Check for incompatible types
    supported_formats = export_matrix[document.doctype]
    if not target_format in supported_formats:
        log.warning(f"cannot export document: '{document.source.name}'") 
        log.warning(f"incompatible format: '{target_format.value}'")
        return None

    # Format destination
    if not destination:
        log.debug('no destination specified, using default')
        desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        destination = Filepath(desktop + '\\' + EXPORT_FOLDER_DEFAULT)

    # Set up valid destination
    if not Path(destination.complete).exists():
        log.debug('destination does not exist, creating new folder')
        os.makedirs(destination.complete)

    # Format output file
    root = document.source.root
    extension = target_format.value
    return Filepath(fr"{destination.complete}\{root}.{extension}")


# OBJECTS
# Which document types are compatible with which export formats?
export_matrix: Dict[SWDocType, List[Type[SWExportFormat]]] = {
    SWDocType.PART: [
        SWExportFormat.IMAGE,
        SWExportFormat.PARASOLID,
        SWExportFormat.STEP,
        SWExportFormat.STL
    ],
    SWDocType.ASSEMBLY: [
        SWExportFormat.IMAGE,
        SWExportFormat.PARASOLID,
        SWExportFormat.STEP,
        SWExportFormat.STL
    ],
    SWDocType.DRAWING: [
        SWExportFormat.DXF,
        SWExportFormat.PDF
    ]
}
