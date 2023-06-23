SolidWrap
=========

*Copyright (c) 2023 Sean Yeatts. All rights reserved.*

This module provides an API for streamlined interfacing with SolidWorks
and Enterprise PDM processes.

Key Features
------------
- Intuitive Pythonic syntax for interacting with SolidWorks models & manipulating PDM states
- B
- Full API documentation detailing the inner-workings of the module.

Quickstart
----------

Key ``import`` statements:

.. code:: python

  import solidwrap                         # top level module import
  from solidwrap import solidworks, vault  # explicit access to SolidWorks and PDM processes

The following methods must be called before you can utilize the SolidWrap API:

.. code:: python

  solidworks.connect(version=2021):  # connect to SolidWorks application
  vault.connect("VAULT_NAME_HERE")   # connect to PDM Vault ( case sensitive )

A simple example script that opens, rebuilds, saves, and closes a .sldprt file:

.. code:: python

  # Performs a simple example operation
  def example():
      file = sw.Filepath(f"C:\My_Vault\Test_Part_01.sldprt")  # create a Filepath object for the target file
      if my_model := solidworks.open(filepath=file)           # if the file opens succesfully...
          solidworks.safeclose(model=my_model)                # ... then safeclose (rebuild, save, close) the file

  # MAIN
  def main():
      if not solidworks.connect(version=2021):  # connect to SW
          vault.connect("My_Vault")             # connect to Vault
          example()

      solidworks.disconnect()                   # terminate SW connection
      vault.disconnect()                        # terminate Vault connection



Installation
------------

*This module is NOT publicly available via PyPI. The contents of this module must be saved locally to a specified location.*

**Prerequisites**

- Python 3.8 or higher is recommended
- pip 23.0 or higher is recommended

**Procedure**

Extract the contents of this module to a safe location *( you can download the ZIP from this project's main GitHub page )*. Open a new terminal and navigate to the top level directory of your project. Run the following command:

.. code:: sh

  py -m pip install DIRECTORY_HERE\dist\solidwrap-0.0.1.tar.gz


- ``DIRECTORY_HERE`` should be replaced with the complete filepath to the folder where you saved the SolidWrap module contents.
- Depending on the version of SolidWrap you've selected, you may have to change ``0.0.1`` to reflect your specific version.
