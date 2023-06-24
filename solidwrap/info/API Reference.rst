API Reference ( SolidWrap )
===========================

*Copyright (c) 2023 Sean Yeatts. All rights reserved.*

The SolidWorks API ( `SW-API <https://help.solidworks.com/2019/English/SolidWorks/sldworks/c_solidworks_api.htm?verRedirect=1>`_ ) and PDM API ( `PDM-API <https://help.solidworks.com/2019/English/api/epdmapi/Welcome-epdmapi.html?id=2a67aaceb6984695a5ce8a75121853f3#Pg0>`_ ) are built on the Component Object Model ( `COM <https://learn.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal>`_ ) to provide an interface with SolidWorks applications. They do not, however, have direct support for Python. SolidWrap leverages its own implementation of the COM pipeline to provide a streamlined Pythonic interface with SW-API and PDM-API.

The core of SolidWrap relies on two objects: ``solidworks`` and ``vault``. These are treated as singletons; they come pre-instanced by the module and should not be manually created by the user. Most interactions with SolidWrap should flow through these objects.

See the Appendix for an overview of the helper classes ( ``Filepath`` & ``Model`` ) that are embedded in many of the SolidWrap class methods.

``solidworks`` ( object )
-------------------------
The core object of the API. It serves as a representation of SolidWorks, and is responsible for handling all SolidWorks commands.

Attributes
``````````
- client ( `ISldWorks <https://help.solidworks.com/2019/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks.html?verRedirect=1>`_ ) - Top level interface for SW-API
- version ( `int <https://www.w3schools.com/python/python_datatypes.asp>`_ ) - Release year of the target SolidWorks application version ( ex. 2021 )

Methods
```````
.. code:: python

  # Establishes a connection to the SolidWorks process.
  def connect(version: int = 2021, visible: bool = False):
    """
    Parameters:
      - version ( int ) - Release year of the target SolidWorks application version (ex. 2021).
      - visible ( bool ) - Whether or not the SolidWorks application window should be displayed.
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
      - filepath ( Filepath ) - Filepath of the target model.
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
  # Saves the target model.
  def save(model: Model):
    """
    Parameters:
      - model ( Model ) - Target Model to save.
    """
    
  # Rebuilds the target model.
  def rebuild(model: Model):
    """
    Parameters:
      - model ( Model ) - Target Model to rebuild.
    """

  # Exports the target model as the prescribed file type.
  def export(model: Model, as_type: str = "PNG"):
    """
    Parameters:
      - model ( Model ) - Target Model to export.
      - as_type ( str ) - Extension of the desired export file type.
    """

  # Freezes the target model's feature tree.
  def freeze(model: Model):
    """
    Parameters:
      - model ( Model ) - Target Model to freeze.
    """

  # Hides / Shows all of the target model's sketches, reference geometry, etc.
  def declutter(model: Model, declutter: bool = True):
    """
    Parameters:
      - model ( Model ) - Target Model to declutter.
      - declutter ( bool ) - Hide / show toggle.
    """

  # Declutters the viewport and orients an isometric model view.
  def stage(model: Model):
    """
    Parameters:
      - model ( Model ) - Target Model to stage.
    """


``vault`` ( object )
--------------------
A representation of the PDM Vault. All PDM interactions ( state changes, checking in / out, etc. ) are handled through this object.

Attributes
``````````
- client ( `IEdmVault5 <https://help.solidworks.com/2019/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmvault5.html?verRedirect=1>`_ ) - Top level interface for PDM-API
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

  # Authenticates login credentials for PDM Vault.
  def authenticate():
    """
    Parameters:
      - NONE
    """

  # Checks out model from PDM Vault.
  def checkout(filepath: Filepath):
    """
    Parameters:
      - filepath ( Filepath ) - Target file to check out.
    """

  # Checks out multiple models from the PDM Vault.
  def batch_checkout(files):
    """
    Parameters:
      - files ( Filepath ) - Target files to check out.
    """

  # Checks in model from PDM Vault.
  def checkin(filepath: Filepath):
    """
    Parameters:
      - filepath ( Filepath ) - Target file to check in.
    """

  # Checks in multiple models from the PDM Vault.
  def batch_checkin(files):
    """
    Parameters:
      - files ( Filepath ) - Target files to check in.
    """

  # Changes model's PDM state to the prescribed value, if allowed.
  def change_state(filepath: Filepath, state: str = "WIP", message: str = "SolidWrap Automated State Change"):
    """
    Parameters:
      - filepath ( Filepath ) - Target file whos state should change.
      - state ( str ) - Literal name of the target state.
      - message ( str ) - Message to include in the file's PDM history.
    """

Appendix
--------
Two container classes are used to simplify the concept of a SolidWorks "document." SW-API tends to prefer the use of complete filepaths as direct references to documents. This is cumbersome, and a less verbose solution is implemented by SolidWrap to streamline file references.

``Model`` ( class )
-------------------
A container that holds Filepath, IModelDoc2, and IEdmFile5 information.

Members
```````
- filepath ( `Filepath <https://github.com/SeanYeatts/QuickPathStr>`_ ) - Filepath representation
- swobj ( `IModelDoc2 <https://help.solidworks.com/2020/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2.html>`_ ) - SW-API representation
- pdmobj ( `IEdmFile5 <https://help.solidworks.com/2019/English/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmFile5.html?verRedirect=1>`_ ) - PDM-API representation [#f]_

``Filepath`` ( class )
----------------------
This class is a simple container that breaks up a complete filepath into its constituent components. It simplifies file references by allowing methods to pass ``Filepath`` objects instead of long, verbose strings. See the `GitHub repository <https://github.com/SeanYeatts/QuickPathStr>`_ for complete details. 

.. rubric::
-----------

.. [#f] `IEdmFile5 <https://help.solidworks.com/2019/English/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmFile5.html?verRedirect=1>`_ data is not yet captured in this release of SolidWrap. Attempting to call this class member will throw an error.
