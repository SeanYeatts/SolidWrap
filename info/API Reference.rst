API Reference
-------------

See the Appendix for an overview of the helper classes ``Filepath`` and ``Model`` that are present in many of the SolidWrap functions.

The core of the API relies on two objects: ``solidworks`` and ``vault``. These are treated as singletons; they come pre-instanced by the module and should not be manually created by the user. Most interactions with the SolidWrap API should flow through these objects.

The ``solidworks`` object represents a connection to the SolidWorks desktop application:

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




The ``vault`` object represents a connection to the PDM Vault:


Appendix
--------

Two container classes are used to simplify the concept of a "model" within the SolidWorks API's preferred file manipulation format.
