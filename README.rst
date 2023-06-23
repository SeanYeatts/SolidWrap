SolidWrap
=========

*Copyright (c) 2023 Sean Yeatts. All rights reserved.*

This module provides an API for streamlined interfacing with SolidWorks
and Enterprise PDM processes.

Key Features
------------
- Intuitive Pythonic syntax for interacting with SolidWorks models & manipulating PDM states
- B
- C

API Reference
-------------
Key ``import`` statements:

.. code:: python

  import solidwrap                         # main SolidWrap module
  from solidwrap import solidworks, vault  # direct access to SolidWorks and PDM processes


Installation
------------
**Python 3.8 or higher is recommended**

*This module is NOT publicly available via PyPI. The contents of this module must be saved locally to a specified location.*

**Step 1**

Extract the contents of this module to a safe location ( you can download the ZIP from this project's main GitHub page ).

**Step 2**

Open a new terminal and navigate to the top level directory of your project. Run the following command:

.. code:: sh

  py -m pip install DIRECTORY\dist\solidwrap-0.0.1.tar.gz

Notes:

- ``DIRECTORY`` should be replaced with the complete filepath to the folder where you saved the SolidWrap module contents.
- Depending on the version of SolidWrap you've selected, you may have to change ``0.0.1`` to reflect your specific version. If you're not sure what version you have, check the ``tar.gz`` file in this module's ``dist`` folder.
