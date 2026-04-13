#Build Advanced_VLookup app
import PyInstaller.__main__
import os
import sys

def build_app():
    # Define absolute paths to avoid path‑related issues
    script_dir = os.path.dirname(os.path.abspath(__file__))
    main_script = os.path.join(script_dir, "excel_suite.py")
    icon_path = os.path.join(script_dir, "icon.ico")
    
    # Check whether icon.ico exists in the directory
    if not os.path.exists(icon_path):
        print(f"❌ Error: File not found at path '{icon_path}'!")
        print("Ensure that a file named icon.ico is present in the same directory as build.py.")
        sys.exit(1)

    # PyInstaller uses ';' on Windows and ':' on macOS/Linux for --add-data
    separator = ';' if os.name == 'nt' else ':'
    
    # PyInstaller arguments
    args = [
        main_script,                         # Source file
        "--name=Advanced_VLookup",           # The .exe name
        "--windowed",                        # Hide the command console (no black window displayed).
        "--onefile",                         # Package everything into a single .exe file.
        "--clean",                           # Clean the cache before building.
        f"--icon={icon_path}",               # Set the icon for the .exe file in Windows Explorer.
        
        # Add the icon to the package so it can be used by the graphical interface
        # Place it inside the internal "assets" folder, as expected by the logic in excel_suite.py
        f"--add-data={icon_path}{separator}assets",
        
        # Ensure that CustomTkinter resources (themes, color schemes) are included
        "--collect-all=customtkinter",
    ]

    print("🚀 Starting the executable build... Please wait, this may take a few minutes.\n")
    
    # Run PyInstaller
    PyInstaller.__main__.run(args)
    
    print("\n✅ Build completed successfully!")
    print(f"You can find the application in the folder: {os.path.join(script_dir, 'dist')}")

if __name__ == "__main__":
    build_app()
