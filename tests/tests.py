'''Copyright (c) 2024 Sean Yeatts. All rights reserved.'''

from __future__ import annotations


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
