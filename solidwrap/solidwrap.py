'''Copyright (c) 2024 Sean Yeatts. All rights reserved.'''

from __future__ import annotations


# IMPORTS - EXTERNAL
from quickpathstr import Filepath           # file / folder manipulation

# IMPORTS - STANDARD LIBRARY
import os                                   # file / folder manipulation
import pythoncom        as pycom            # used with win32com.client
import subprocess       as subproc          # process termination
import win32com.client  as win              # COM object handling
from pathlib import Path                    # file / folder manipulation
from typing import Any, Dict, List, Type    # type checking

# IMPORTS - PROJECT
from solidwrap.containers import SWAuthState, SWDocType, SWExportFormat, SWDocument
from solidwrap.logger import log
from solidwrap.utilities import compute_client_key, singleton
from solidwrap.utilities import (
    AUTOMATION_MESSAGE,
    EXPORT_FOLDER_DEFAULT,
    SUBPROCESS_NAME,
    VAULT_DISPATCH_KEY
)


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
    Wrapper for the SolidWorks API. Represents a SolidWorks client.
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

                # Enforce visibility
                SolidWorks.client.visible               = not headless
                SolidWorks.client.UserControlBackground = headless
                SolidWorks.client.Frame.KeepInvisible   = headless
                
                log.info('connection successfully established')
                return True
        except Exception as exception:
            log.error('failed to establish connection')
            log.error(exception)
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
            log.error('failed to disconnect from SolidWorks client')
            log.error(exception)
            return False

    # DOCUMENT MANAGEMENT METHODS
    def open(self, filepath: Filepath) -> SWDocument:
        '''Opens a document using a prescribed filepath.'''
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
        '''Closes a target document ( WITHOUT rebuild & save operations ).'''
        log.info(f"closing document: '{document.source.name}'")

        # SolidWorks API call
        SolidWorks.client.CloseDoc(document.source.complete)

    def safeclose(self, document: SWDocument) -> None:
        '''Closes a target document ( WITH rebuild & save operations ).'''
        self.rebuild(document, False)
        self.save(document)
        self.close(document)

    def save(self, document: SWDocument) -> None:
        '''Saves a target document.'''
        log.info(f"saving document: '{document.source.name}'")

        # Define COM VARIANT args
        options     = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)
        errors      = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)
        warnings    = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, 1)

        # SolidWorks API call
        document.swobj.Save3(options, errors, warnings)

    def rebuild(self, document: SWDocument, top_only: bool = False) -> None:
        '''Rebuilds a target document.'''
        log.info(f"rebuilding document: '{document.source.name}'")
        
        # Define COM VARIANT args
        top_only = win.VARIANT(pycom.VT_BYREF | pycom.VT_I4, top_only)

        # Execute SW-API method
        document.swobj.ForceRebuild3(top_only)

    # HIGH-LEVEL METHODS
    def export(self, document: SWDocument, as_type: SWExportFormat,
        destination: Filepath = None) -> None:
        '''Exports a document using a prescribed format.'''
        
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
        '''Declutters the viewport and orients an isometric model view.'''
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
        '''Freezes a target document's Feature Tree.'''
        log.info(f"freezing document: '{document.source.name}'")

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
        '''Gets the last feature in a document's Feature Tree.'''
        log.debug('getting last feature in Feature Tree')
        
        # Gather necessary components of the Feature Tree
        manager = document.swobj.FeatureManager     # get Feature Manager
        tree = manager.GetFeatureTreeRootItem2(0)   # get Feature Tree root
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
    Wrapper for the SolidWorks API. Represents a PDM Vault client.
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
            log.error('failed to establish connection')
            log.error(exception)
            return False

    def disconnect(self) -> bool:
        '''Terminates the PDM Vault connection.'''
        log.critical(f"disconnecting from PDM Vault ( {self.name} )")

        # Attempt to disconnect from the process
        try:
            pycom.CoUninitialize()
            Vault.client = None
            self.authorized = SWAuthState.UNAUTHORIZED
            log.info('connection successfully terminated')
            return True
        except Exception as exception:
            log.error('failed to disconnect from PDM Vault client')
            log.error(exception)
            return False

    def authorize(self) -> bool:
        '''Authenticates login credentials for the PDM Vault.'''
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
            log.error('failed authenticate login credentials')
            log.error(exception)
            return False

    # FILE STATE METHODS
    def checkin(self, filepath: Filepath, comment: str = None) -> None:
        """
        Checks in a document to the PDM Vault.
        
        IMPORTANT: this method will not work if the file is currently
        open as a document in SolidWorks. Close the file before use.
        """
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
            log.info(f"file is already checked in: '{filepath.name}'")
            return None
        
        # Execute PDM-API method    
        file.UnlockFile(0, message)
            
    def checkout(self, filepath: Filepath) -> None:
        """
        Checks out a document from the PDM Vault.
        
        IMPORTANT: this method will not work if the file is currently
        open as a document in SolidWorks. Close the file before use.
        """
        log.info(f"checking out file: '{filepath.name}'")

        # Get PDM-API objects
        directory = Vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                         # IEdmFile

        # Check current document state
        if file.IsLocked:
            log.info(f"file is already checked out: '{filepath.name}'")
            return None
        
        # Execute PDM-API method
        file.LockFile(directory.ID, 0)

    def undo_checkout(self, filepath: Filepath) -> None:
        """
        Reverts a check out from the PDM Vault.
        
        IMPORTANT: this method will not work if the file is currently
        open as a document in SolidWorks. Close the file before use.
        """
        log.info(f"undoing check out: '{filepath.name}'")

        # Get PDM-API objects
        directory = Vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                         # IEdmFile

        # Check current document state
        if not file.IsLocked:
            log.info(f"file is not checked out: '{filepath.name}'")
            return None
        
        # Execute PDM-API method
        file.UndoLockFile(0)

    # WORKFLOW STATE METHODS
    def change_state(self, filepath: Filepath, transition: str, comment: str = None) -> None:
        '''Changes the PDM state of a file using a provided transition.'''

        # Get PDM-API objects
        directory = Vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                         # IEdmFile
        
        # Get all possible transitions
        index: int = 0
        transitions: List = []
        current = file.CurrentState
        position = current.GetFirstTransitionPosition()
        while not position.IsNull:
            possibility = current.GetNextTransition(position)
            transitions.append(possibility)
            index += 1
        
        # Get next state by checking against possible transitions
        found: bool = False
        for test in transitions:
            if test.Name == transition:
                next_state = test.ToState
                log.info(fr"transitioning file ( '{filepath.name}' ) to state: {next_state.Name}")
                found = True
        if not found:
            log.warning(f"failed to execute transition: '{transition}'")
            return
        
        # Format a default comment
        if comment is None:
            comment = AUTOMATION_MESSAGE

        # Define COM VARIANT args
        arg1 = win.VARIANT(pycom.VT_I4, int(directory.ID))
        arg2 = win.VARIANT(pycom.VT_BSTR, comment)

        # SolidWorks API call
        file.ChangeState(next_state, arg1.value, arg2.value, 0)

    # DATA RETRIEVAL METHODS
    def get_revision(self, filepath: Filepath) -> str:
        '''Gets a file's release version.'''
        log.debug('release revision')

        # Get PDM-API objects
        directory = Vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                         # IEdmFile
        return file.CurrentRevision
    
    def get_state(self, filepath: Filepath) -> str:
        '''Gets a file's current PDM state.'''
        
        # Get PDM-API objects
        directory = Vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                         # IEdmFile
        return file.CurrentState.Name

    def get_transitions(self, filepath: Filepath) -> List[str]:
        '''Gets all of the possible state transitions for a file's current
        PDM state.'''
        
        # Get PDM-API objects
        directory = Vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                         # IEdmFile
        
        index: int = 0
        transitions: List = []
        current = file.CurrentState
        position = current.GetFirstTransitionPosition()
        while not position.IsNull:
            transition = current.GetNextTransition(position)
            transitions.append(transition.Name)
            index += 1
        return transitions

    def get_checkout_user(self, filepath: Filepath) -> Any:
        
        # Get PDM-API objects
        directory = Vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                         # IEdmFile
        if file.IsLocked:
            return file.LockedByUser.Name

    def get_configurations(self, filepath: Filepath) -> List[str]:

        # Get PDM-API objects
        directory = Vault.client.GetFolderFromPath(filepath.directory)  # IEdmFolder
        file = directory.GetFile(filepath.name)                         # IEdmFile
        configs = file.GetConfigurations()

        index: int = 0
        results: List = []
        position = configs.GetHeadPosition()
        while not position.IsNull:
            value = configs.GetNext(position)
            results.append(value)
            index += 1
        if '@' in results:
            results.remove('@')
        return results

# FUNCTIONS
def prepare_export(document: SWDocument, target_format: SWExportFormat,
        destination: Filepath = None) -> Filepath:
    '''Prepares a document for an export operation.'''
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
export_matrix: Dict[SWDocType, List[SWExportFormat]] = {
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
