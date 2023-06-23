API Reference
=============
The core of the API relies on two objects: ``solidworks`` and ``vault``. These are treated as singletons; they come pre-instanced by the module and should not be manually created by the user. Most interactions with the SolidWrap API should flow through these objects.

See the Appendix for an overview of the helper classes ``Filepath`` and ``Model`` that are present in many of the SolidWrap functions.

The ``solidworks`` Object
-------------------------
Attributes
``````````
- client ( `COM <https://learn.microsoft.com/en-us/windows/win32/com/the-component-object-model>`_ ) - Direct reference to the SolidWorks application
- version ( `int <https://www.w3schools.com/python/python_datatypes.asp>`_ ) - Release year of the target SolidWorks application version (ex. 2021)

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



The ``vault`` Object
--------------------
Attributes
``````````
- client ( `IEdmVault5 <https://help.solidworks.com/2019/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmvault5.html?verRedirect=1>`_ ) - Direct reference to the PDM Vault
- name ( `str <https://www.w3schools.com/python/python_datatypes.asp>`_ ) - Literal name of the PDM Vault
- auth_state ( `bool <https://www.w3schools.com/python/python_datatypes.asp>`_ ) - Authorization flag; indicates successful login credentials

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
Two container classes are used to simplify the concept of a SolidWorks "model." The SolidWorks API tends to prefer the use of complete filepaths as direct references to documents. This is cumbersome, and a less verbose solution is imlpemented by the SolidWrap API to compartmentalize file references.

``Model``
---------
Attributes
``````````
- Filepath ( `Filepath <https://github.com/SeanYeatts/QuickPathStr>`_ ) - Filepath representation of the Model
- swobj ( `IModelDoc2 <https://help.solidworks.com/2020/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2.html>`_ ) - SolidWorks API representation of the Model

