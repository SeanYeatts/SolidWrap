from solidwrap import solidworks, vault
from quickpathstr import Filepath

def execute():
    """EXAMPLE EXECUTION LOGIC"""
    
    files = [
        Filepath(fr"C:\Goddard_Vault\Users\SYeatts\Scripts\Test_Part_01.SLDPRT"),
        Filepath(fr"C:\Goddard_Vault\Users\SYeatts\Scripts\Test_Part_02.SLDPRT"),
        Filepath(fr"C:\Goddard_Vault\Users\SYeatts\Scripts\Test_Part_03.SLDPRT"),
    ]
    
    # Example export workflow for generating Agile attachments

    # Clean up feature trees
    vault.batch_checkout(files)
    for file in files:
        if model := solidworks.open(filepath=file):
            solidworks.freeze(model)
            solidworks.safeclose(model=model)
    vault.batch_checkin(files)

    # Export Agile attachments
    for file in files:
        if model := solidworks.open(filepath=file):
            solidworks.export(model=model, as_type="png")
            solidworks.export(model=model, as_type="x_t")
            solidworks.close(model=model)

# MAIN ENTRYPOINT
def main():
    if not solidworks.connect(version=2021):
        vault.connect("Goddard_Vault")
        execute()
    input("\nPress any key to continue...")

# TOP LEVEL SCRIPT ENTRYPOINT
if __name__ == '__main__':
    main()
