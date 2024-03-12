API Reference ( SolidWrap )
===========================

*Copyright (c) 2024 Sean Yeatts. All rights reserved.*

The SolidWorks API ( `SW-API <https://help.solidworks.com/2019/English/SolidWorks/sldworks/c_solidworks_api.htm?verRedirect=1>`_ ) and PDM API ( `PDM-API <https://help.solidworks.com/2019/English/api/epdmapi/Welcome-epdmapi.html?id=2a67aaceb6984695a5ce8a75121853f3#Pg0>`_ ) are built on the Component Object Model ( `COM <https://learn.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal>`_ ) to provide an interface with SolidWorks applications. They do not, however, have direct support for Python. SolidWrap leverages its own implementation of the COM pipeline to provide an easy-to-use, streamlined Pythonic interface with SW-API and PDM-API.

----

The core of SolidWrap relies on two classes: ``SolidWorks`` and ``Vault``. These are treated as singletons; they must be instanced by the user prior to making any SolidWrap API calls. All interactions with SolidWorks flows through these objects.

See the Appendix for an overview of the helper classes ( ``Filepath`` & ``SWDocument`` ) that are embedded in many of the SolidWrap class methods.

**NOTE : Items labelled 'WIP' are in development and are not guaranteed to function as expected. They may likely even fail entirely.**

----

``SolidWorks`` ( Class )
------------------------
The core class of the API. It serves as a representation of SolidWorks, and is responsible for handling all SolidWorks commands.

Attributes
``````````
- client ( `ISldWorks <https://help.solidworks.com/2019/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks.html?verRedirect=1>`_ ) - Top level interface for SW-API
- version ( `int <https://www.w3schools.com/python/python_datatypes.asp>`_ ) - Release year of the target SolidWorks application version ( ex. 2021 )

Methods
```````
.. code:: python

  # Creates a connection to the SolidWorks client.
  def connect(headless: bool = False) -> bool:
    """
    Params:
      - version ( int ) - application version, denoted by release year
      - headless ( bool ) - hides the SolidWorks application window
    Returns:
      - ( bool ) - True if successful
    """

  # Terminates the SolidWorks connection.
  def disconnect(silent: bool = True) -> bool:
    """
    Params:
      - silent ( bool ) - keeps the SolidWorks application open
    Returns:
      - ( bool ) - True if successful
    """

  # Opens a document using a prescribed filepath.
  def open(filepath: Filepath) -> SWDocument:
    """
    Params:
      - filepath ( Filepath ) - Filepath of the target document
    Returns:
      - ( SWDocument ) - resulting document object
    """

  # Closes a target document ( WITHOUT rebuild & save operations ).
  def close(document: SWDocument) -> None:
    """
    Params:
      - document ( SWDocument ) - target document
    """

  # Closes a target document ( WITH rebuild & save operations ).
  def safeclose(document: SWDocument) -> None:
    """
    Params:
      - document ( SWDocument ) - target document
    """

  # Saves a target document.
  def save(document: SWDocument) -> None:
    """
    Params:
      - document ( SWDocument ) - target document
    """
    
  # Rebuilds a target document.
  def rebuild(document: SWDocument, top_only: bool = False) -> None:
    """
    Params:
      - document ( SWDocument ) - target document
      - top_only ( bool ) - restricts operation to top level document
    """

  # Exports a document using a prescribed format.
  def export(document: SWDocument, as_type: SWExportFormat,
        destination: Filepath = None) -> None:
    """
    Params:
      - document ( SWDocument ) - target document
      - as_type ( SWExportFormat ) - describes an export format
      - destination ( Filepath ) - assigns an output directory
    """

  # Declutters the viewport and orients an isometric model view.
  def stage(document: SWDocument) -> None:
    """
    Params:
      - document ( SWDocument ) - target document
    """

  # Freezes a target document's Feature Tree.
  def freeze(document: SWDocument) -> None:
    """
    Params:
      - document ( SWDocument ) - target document
    """


``Vault`` ( Class )
--------------------
A representation of the PDM Vault. All PDM interactions ( state changes, checking in / out, etc. ) are handled through this object.

Attributes
``````````
- client ( `IEdmVault5 <https://help.solidworks.com/2019/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmvault5.html?verRedirect=1>`_ ) - Top level interface for PDM-API
- name ( `str <https://www.w3schools.com/python/python_datatypes.asp>`_ ) - Literal name of the PDM Vault
- authorized ( SWAuthState ) - Authorization flag indicating successful login credentials

Methods
```````
.. code:: python

  # Creates a connection to the PDM Vault client.
  def connect() -> bool:
    """
    Returns:
      - ( bool ) - True if successful
    """

  # Terminates the PDM Vault connection.
  def disconnect() -> bool:
    """
    Returns:
      - ( bool ) - True if successful
    """

  # Authenticates login credentials for the PDM Vault.
  def authorize() -> bool:
    """
    Returns:
      - ( bool ) - True if successful
    """

  # Checks in a document to the PDM Vault.
  def checkin(filepath: Filepath, comment: str = None) -> None:
    """
    Params:
      - filepath ( Filepath ) - Filepath of the target document
      - comment ( str ) - message to include for check in history
    """

  # Checks out a document from the PDM Vault.
  def checkout(filepath: Filepath) -> None:
    """
    Params:
      - filepath ( Filepath ) - Filepath of the target document
    """

  # Reverts a check out from the PDM Vault.
  def undo_checkout(filepath: Filepath) -> None:
    """
    Params:
      - filepath ( Filepath ) - Filepath of the target document
    """

----

Appendix
--------

I. Containers
`````````````
A number of container classes are used to simplify various concepts within the context of file / document management. An overview of these classes is outlined below:

``SWDocument`` ( Class )
-------------------
A wrapper for Filepath, IModelDoc2, and IEdmFile5 information.

Members
```````
- source ( `Filepath <https://github.com/SeanYeatts/QuickPathStr>`_ ) - Filepath representation
- swobj ( `IModelDoc2 <https://help.solidworks.com/2020/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2.html>`_ ) - SW-API representation
- pdmobj ( `IEdmFile5 <https://help.solidworks.com/2019/English/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmFile5.html?verRedirect=1>`_ ) - PDM-API representation [#f]_

``Filepath`` ( Class )
----------------------
This class is a simple container that breaks up a complete filepath into its constituent components. It simplifies file references by allowing methods to pass ``Filepath`` objects instead of long, verbose strings. See the `GitHub repository <https://github.com/SeanYeatts/QuickPathStr>`_ for complete details. 


II. Multithreading
``````````````````
A specific paradigm must be followed in multithreaded environments to ensure that SolidWrap functions as intended. This is a consequence of the module's reliance on COM objects, which do not inherently support satisfactory inter-thread communication behavior.

This means that for GUI applications, SolidWrap objects should **NOT** be instantiated on the main GUI thread. It is recommended that a dedicated thread be allocated for the ``SolidWorks`` and ``Vault`` objects. A command queue can be implemented to pass method calls to the dedicated thread. Similarly, a results queue can be implemented to act on results generated by method calls.

.. rubric::
-----------

.. [#f] `IEdmFile5 <https://help.solidworks.com/2019/English/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmFile5.html?verRedirect=1>`_ data is not yet captured in this release of SolidWrap. Attempting to call this class member will throw an error.
