import solidwrap as sw
from solidwrap import solidworks, vault
from solidwrap import extensions

def execute():
    """EXAMPLE EXECUTION LOGIC"""
    
    # Simple hard-coded test files
    files = [
        sw.Filepath(fr"C:\{vault.name}\Test_Part_01.SLDPRT"),
        sw.Filepath(fr"C:\{vault.name}\Test_Part_02.SLDPRT"),
        sw.Filepath(fr"C:\{vault.name}\Test_Part_03.SLDPRT"),
    ]
    
    # Example export workflow
    # -----------------------
    # Clean up feature trees
    vault.batch_checkout(files)
    for file in files:
        if model := solidworks.open(filepath=file):
            solidworks.freeze(model)
            solidworks.safeclose(model=model)
    vault.batch_checkin(files)

    # Export files
    for file in files:
        if model := solidworks.open(filepath=file):
            solidworks.export(model=model, as_type="png")
            solidworks.export(model=model, as_type="x_t")
            solidworks.close(model=model)

# MAIN ENTRYPOINT
def main():
    if not solidworks.connect(version=2021):
        vault.connect('My_Vault')
        execute()
    input("\nPress any key to continue...")
    solidworks.disconnect()

# TOP LEVEL SCRIPT ENTRYPOINT
if __name__ == '__main__':
    main()
