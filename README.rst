SolidWrap
=========

*Copyright (c) 2024 Sean Yeatts. All rights reserved.*

This module wraps the SolidWorks and PDM APIs for a streamlined, Pythonic workflow.

Key Features
------------
- Intuitive Pythonic syntax for interacting with SolidWorks documents & PDM files.
- Timestamped logging for insights into automated task performance.
- Full API documentation detailing module components.

Navigate to the `API Reference <https://github.com/SeanYeatts/SolidWrap/blob/main/info/API%20Reference.rst>`_ for a complete walkthrough of the module.

Quickstart
----------

Key ``import`` statements :

.. code:: python

  import solidwrap                         # top level module import
  from solidwrap import SolidWorks, Vault  # access to SolidWorks and PDM processes
  from solidwrap import Filepath           # generic container for file objects

The following methods must be called before you can utilize the SolidWrap API :

.. code:: python

  solidworks = SolidWorks(version=2023)   # instantiate SolidWorks
  vault = Vault(name='VAULT-NAME')        # instantiate PDM Vault ( case sensitive )

  solidworks.connect()   # connect to SolidWorks client
  vault.connect()        # connect to PDM Vault client

Example - a simple script that exports a group of '.SLDPRT' documents to various formats :

.. code:: python

  # IMPORTS
  from solidwrap import SolidWorks, Vault         # core classes
  from solidwrap import Filepath, SWExportFormat  # containers
  
  
  # FUNCTIONS
  def export_example(solidworks: SolidWorks, vault: Vault) -> None:
  
      # Set up some example files
      files = [
          Filepath(fr"C:\{vault.name}\Test_Part_01.SLDPRT"),
          Filepath(fr"C:\{vault.name}\Test_Part_02.SLDPRT"),
          Filepath(fr"C:\{vault.name}\Test_Part_03.SLDPRT")
      ]
  
      # Export a variety of formats
      for file in files:
          if (document := solidworks.open(file)):
              solidworks.export(document, SWExportFormat.IMAGE)      # .png
              solidworks.export(document, SWExportFormat.PARASOLID)  # .x_t
              solidworks.close(document)
  
  
  # MAIN DEFINITION
  def main() -> None:
      
      # Instantiate core objects
      solidworks = SolidWorks(2023)
      vault = Vault('My-Vault')
  
      # Initialize connections
      if not solidworks.connect(headless=False):
          return None
      if not vault.connect():
          return None
      
      # Do some stuff...
      export_example(solidworks, vault)
      
      # Terminate connections
      solidworks.disconnect(silent=False)
      vault.disconnect()
  
  
  # TOP LEVEL ENTRY POINT
  if __name__ == '__main__':
      main()


Installation
------------
**Prerequisites**

- Python 3.8 or higher is recommended
- pip 23.0 or higher is recommended

**For a pip installation**

Open a new Command Prompt. Run the following command:

.. code:: sh

  py -m pip install solidwrap

**For a local installation**

Extract the contents of this module to a safe location. Open a new terminal and navigate to the top level directory of your project. Run the following command:

.. code:: sh

  py -m pip install "DIRECTORY_HERE\solidwrap\dist\solidwrap-2.0.0.tar.gz"

- ``DIRECTORY_HERE`` should be replaced with the complete filepath to the folder where you saved the SolidWrap module contents.
- Depending on the release of SolidWrap you've chosen, you may have to change ``2.0.0`` to reflect your specific version.
