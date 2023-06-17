from solidwrap      import SolidWrap
from quickpathstr   import Filepath

# Example
def execute(client: SolidWrap):
    """EXAMPLE CLIENT LOGIC"""
    
    test_file = Filepath(fr"C:\Users\seany\Desktop\Test Vault\TASRS\Test_Part_02.SLDPRT")
    model = client.open_model(test_file)

    input("\nPress any key to continue...")
    client.close(model)

    input("\nPress any key to continue...")
    client.disconnect() # only called explicitly because we set debug=True

# MAIN ENTRYPOINT
def main():
    # Core object instantiation
    if application := SolidWrap(version=2019, debug=True):
        execute(application)

# TOP LEVEL SCRIPT ENTRYPOINT
if __name__ == '__main__':
    main()
