"""
Copyright (c) 2024 Sean Yeatts. All rights reserved.

A collection of classes for convenient handling of logically related
data types.
"""

from __future__ import annotations


# IMPORTS - EXTERNAL
from quickpathstr import Filepath

# IMPORTS - STANDARD LIBRARY
import os
from enum import Enum, auto
from typing import Any

# IMPORTS - PROJECT
from solidwrap.utilities import FileSize


# ENUMS
class SWAuthState(Enum):
    '''Logical wrapper for login states.'''
    AUTHORIZED = auto()
    UNAUTHORIZED = auto()


class SWDocType(Enum):
    '''Logical wrapper for document types.'''
    PART        = 'SLDPRT'
    ASSEMBLY    = 'SLDASM'
    DRAWING     = 'SLDDRW'


class SWExportFormat(Enum):
    '''Collection of supported export formats.'''
    DXF         = 'dxf'
    IMAGE       = 'png'
    PARASOLID   = 'x_t'
    PDF         = 'pdf'
    STEP        = 'step'
    STL         = 'stl'


# CLASSES
class SWDocument:
    '''Logical wrapper for SolidWorks & PDM API objects.'''

    # FUNDAMENTAL METHODS
    def __init__(self, swobj: Any) -> None:

        # SolidWorks API COM object
        self.swobj: Any = swobj # IModelDoc2

        # Core attributes
        self.source     = Filepath(swobj.GetPathName)
        self.doctype    = format_doctype(self.source)
        self.size       = FileSize(os.path.getsize(self.source.complete))


# FUNCTIONS
def format_doctype(filepath: Filepath) -> SWDocType:
    '''Formats a doctype from a file's extension.'''
    extension = filepath.extension.replace('.','')
    for string in SWDocType:
        if string.value == extension:
            return string
