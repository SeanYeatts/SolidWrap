API Reference ( SolidWrap )
===========================
Introduction
------------
The `SolidWorks API <https://help.solidworks.com/2019/English/SolidWorks/sldworks/c_solidworks_api.htm?verRedirect=1>`_ ( hereon referred to as 'SW-API' ) is built on the Component Object Model ( `COM <https://learn.microsoft.com/en-us/windows/win32/com/the-component-object-model>`_ ) to provide an interface with SolidWorks software. It does not, however, have direct support for Python. SolidWrap leverages the COM pipeline with its own implementation to provide a Pythonic API for streamlined SolidWorks interfacing.

The core of the API relies on two objects: ``solidworks`` and ``vault``. These are treated as singletons; they come pre-instanced by the module and should not be manually created by the user. Most interactions with the SolidWrap API should flow through these objects.

See the Appendix for an overview of the helper classes ( ``Filepath`` & ``Model`` ) that are embedded in many of the SolidWrap class methods.

``solidworks`` ( object )
-------------------------
The core object of the API. It serves as a representation of SolidWorks, and is responsible for handling all SolidWorks commands.

Attributes
``````````
- client ( `ISldWorks <https://help.solidworks.com/2019/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks.html?verRedirect=1>`_ ) - Top level interface for the SolidWorks API
- version ( `int <https://www.w3schools.com/python/python_datatypes.asp>`_ ) - Release year of the target SolidWorks application version ( ex. 2021 )

Methods
```````
.. code:: python

  # Establishes a connection to the SolidWorks process.
  def connect(version: int = 2021, visible: bool = False):
    """
    Parameters:
      - version ( int ) - Release year of the target SolidWorks application version (ex. 2021)
      - visible ( bool ) - Whether or not the SolidWorks application window should be displayed
    """

  # Terminates the SolidWorks process.
  def disconnect():
    """
    Parameters:
      - None
    """

  # Opens a Part ( .sldprt ) or Assembly ( .sldasm ) file.
  def open(filepath: Filepath) -> Model:
    """
    Parameters:
      - filepath ( Filepath ) - Filepath of the target model
    Returns:
      - ( Model ) - Result
    """

  # Closes the target model ( WITH rebuild & save operations ).
  def safeclose(model: Model):
    """
    Parameters:
      - model ( Model ) - Target Model to close.
    """

  # Closes the target model ( WITHOUT rebuild & save operations ).
  def close(model: Model):
    """
    Parameters:
      - model ( Model ) - Target Model to close.
    """



``vault`` ( object )
--------------------
A representation of the PDM Vault. All PDM interactions ( state changes, checking in / out, etc. ) are handled through this object.

Attributes
``````````
- client ( `IEdmVault5 <https://help.solidworks.com/2019/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmvault5.html?verRedirect=1>`_ ) - Top level interface for the PDM API
- name ( `str <https://www.w3schools.com/python/python_datatypes.asp>`_ ) - Literal name of the PDM Vault
- auth_state ( `bool <https://www.w3schools.com/python/python_datatypes.asp>`_ ) - Authorization flag indicating successful login credentials

Methods
```````
.. code:: python

  # Establishes a connection to the PDM Vault.
  def connect(name: str = "VAULT_NAME_HERE"):
    """
    Parameters:
      - name ( str ) - Literal name of the target PDM Vault
    """

Appendix
--------
Two container classes are used to simplify the concept of a SolidWorks "document." The SolidWorks API tends to prefer the use of complete filepaths as direct references to documents. This is cumbersome, and a less verbose solution is implemented by the SolidWrap API to streamline file references.

``Model`` ( class )
-------------------
A container that holds Filepath, IModelDoc2, and IEdmFile5 information. [#f1]_

Attributes
``````````
- filepath ( `Filepath <https://github.com/SeanYeatts/QuickPathStr>`_ ) - Filepath representation of the Model
- swobj ( `IModelDoc2 <https://help.solidworks.com/2020/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2.html>`_ ) - SolidWorks API representation of the Model

``Filepath`` ( class )
----------------------
This class is a simple container that breaks up a complete filepath into its constituent components. It simplifies file references by allowing methods to pass ``Filepath`` objects instead of long, verbose strings. See the `GitHub repository <https://github.com/SeanYeatts/QuickPathStr>`_ for complete details. 

.. rubric::
-----------

.. [#f1] `IEdmFile5 <https://help.solidworks.com/2019/English/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmFile5.html?verRedirect=1>`_ data is not yet captured in this release of SolidWrap.
