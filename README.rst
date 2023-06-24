SolidWrap
=========

*Copyright (c) 2023 Sean Yeatts. All rights reserved.*

This module wraps the SolidWorks and PDM APIs for a streamlined, Pythonic workflow.

Key Features
------------
- Intuitive Pythonic syntax for interacting with SolidWorks models & PDM states.
- Logging and report generation for profiling automation performance.
- Full API documentation detailing module components.

Navigate to the `API Reference <https://github.com/SeanYeatts/SolidWrap/blob/main/info/API%20Reference.rst#api-reference--solidwrap->`_ for a complete walkthrough of the module.

Quickstart
----------

Key ``import`` statements :

.. code:: python

  import solidwrap                         # top level module import
  from solidwrap import solidworks, vault  # explicit access to SolidWorks and PDM processes

The following methods must be called before you can utilize the SolidWrap API :

.. code:: python

  solidworks.connect(version=2021):  # connect to SolidWorks application
  vault.connect("VAULT_NAME_HERE")   # connect to PDM Vault ( case sensitive )

Example - a simple script that opens, rebuilds, saves, and closes a '.sldprt' file :

.. code:: python

  # Performs a simple example operation
  def example():
      file = sw.Filepath(f"C:\My_Vault\Test_Part_01.sldprt")  # create a Filepath object for the target file

      vault.checkout(file)                                    # check out
      if my_model := solidworks.open(file)                    # if the file opens succesfully...
          solidworks.safeclose(my_model)                      # ... then safeclose ( rebuild, save, close )
      vault.checkin(file)                                     # check in

  # MAIN
  def main():
      if not solidworks.connect(2021):                        # connect to SW
          vault.connect("My_Vault")                           # connect to Vault
          example()                                           # do something

      # Only include these statements if you truly want to quit both processes; they are not mandatory.
      solidworks.disconnect()                                 # terminate SW connection
      vault.disconnect()                                      # terminate Vault connection



Installation
------------

*This module is NOT publicly available via PyPI. The contents of this module must be saved locally to a specified location.*

**Prerequisites**

- Python 3.8 or higher is recommended
- pip 23.0 or higher is recommended

**Procedure**

Extract the contents of this module to a safe location. Open a new terminal and navigate to the top level directory of your project. Run the following command:

.. code:: sh

  py -m pip install "DIRECTORY_HERE\solidwrap\dist\solidwrap-0.0.1.tar.gz"


- ``DIRECTORY_HERE`` should be replaced with the complete filepath to the folder where you saved the SolidWrap module contents.
- Depending on the version of SolidWrap you've chosen, you may have to change ``0.0.1`` to reflect your specific version.
