# Copyright (c) 2023 Sean Yeatts. All rights reserved.
#
# PROJECT:      SolidWrap TEST
# VERSION:      0.0.1
# DESCRIPTION:  Test script for SolidWrap module.
# AUTHOR:       Sean Yeatts
# MODIFIED:     06/20/2023

# MODULE IMPORTS
from solidwrap      import app, vault
from quickpathstr   import Filepath

def do_something(file):
    if model := app.open_model(file):
        app.rebuild(model)
        app.save(model)
        app.force_close(model)

def execute():
    """EXAMPLE APPLICATION LOGIC"""

    test = [Filepath(fr"C:\Goddard_Vault\Users\SYeatts\Scripts\Test_Part_01.SLDPRT"),
            Filepath(fr"C:\Goddard_Vault\Users\SYeatts\Scripts\Test_Part_02.SLDPRT"),
            Filepath(fr"C:\Goddard_Vault\Users\SYeatts\Scripts\Test_Part_03.SLDPRT")
            ]

    # Batch checkout files
    for file in test:
        vault.checkout(file)

    # Batch operate on files
    for file in test:
        do_something(file)

    # Batch checkin files
    for file in test:
        vault.checkin(file)


# MAIN ENTRYPOINT
def main():
    # Connect to SolidWorks process
    if not app.connect(version=2021, debug=True):
        vault.connect("Goddard_Vault")  # specify target PDM Vault
        execute()

    input("\nPress any key to continue...")

# TOP LEVEL SCRIPT ENTRYPOINT
if __name__ == '__main__':
    main()
