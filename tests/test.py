from solidwrap import solidworks, vault
from quickpathstr import Filepath

def execute():
    """EXAMPLE EXECUTION LOGIC"""
    
    file = Filepath(fr"C:\Goddard_Vault\Users\SYeatts\Scripts\Test_Part_02.SLDPRT")
    
    vault.checkout(filepath=file)
    model = solidworks.open(filepath=file)
    solidworks.safeclose(model=model)
    vault.checkin(file)

# MAIN ENTRYPOINT
def main():
    if not solidworks.connect(version=2021):
        vault.connect("Goddard_Vault")
        execute()
    input("\nPress any key to continue...")

# TOP LEVEL SCRIPT ENTRYPOINT
if __name__ == '__main__':
    main()
