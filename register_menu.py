import sys
import os
import winreg
import glob

def find_dll():
    # 1. Check the path we already discovered
    preset_path = r"C:\Program Files (x86)\GTONE\iForm\ImportLib\FilePathCheckerModuleExample.dll"
    if os.path.exists(preset_path):
        return preset_path
    
    # 2. Search in Program Files and Program Files (x86)
    search_dirs = [
        os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
        os.environ.get("ProgramFiles", r"C:\Program Files")
    ]
    for search_dir in search_dirs:
        if os.path.exists(search_dir):
            # Use glob to find the DLL
            pattern = os.path.join(search_dir, "**", "FilePathCheckerModuleExample.dll")
            try:
                matches = glob.glob(pattern, recursive=True)
                if matches:
                    return matches[0]
            except Exception:
                pass
                
    return None

def register():
    python_exe = sys.executable
    pythonw_exe = python_exe.replace("python.exe", "pythonw.exe")
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    converter_py = os.path.join(current_dir, "converter.py")
    
    if not os.path.exists(converter_py):
        print(f"Error: converter.py not found at {converter_py}")
        return False
        
    print(f"Pythonw path: {pythonw_exe}")
    print(f"Converter path: {converter_py}")
    
    # 1. Register Context Menu
    extensions = [".hwp", ".hwpx"]
    command_str = f'"{pythonw_exe}" "{converter_py}" "%1"'
    
    for ext in extensions:
        # Base key
        key_path = f"Software\\Classes\\SystemFileAssociations\\{ext}\\shell\\ConvertToPDF"
        try:
            # Create or open key
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "PDF로 변환 (&P)")
            # Add an icon if possible (using shell32.dll standard icon or a general pdf icon)
            # Shell32.dll index 268 is a nice document/pdf search style icon on Windows
            try:
                winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "shell32.dll,268")
            except Exception:
                pass
            winreg.CloseKey(key)
            
            # Create command subkey
            cmd_key_path = f"{key_path}\\command"
            cmd_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, cmd_key_path)
            winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, command_str)
            winreg.CloseKey(cmd_key)
            
            print(f"Registered context menu for {ext}")
        except Exception as e:
            print(f"Failed to register context menu for {ext}: {e}")
            return False
            
    # 2. Register Security Module DLL
    dll_path = find_dll()
    if dll_path:
        print(f"Found security DLL at: {dll_path}")
        try:
            modules_path = r"SOFTWARE\HNC\HwpAutomation\Modules"
            modules_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, modules_path)
            winreg.SetValueEx(modules_key, "FilePathCheckerModuleExample", 0, winreg.REG_SZ, dll_path)
            winreg.CloseKey(modules_key)
            print("Successfully registered Hancom Security Module DLL in registry.")
        except Exception as e:
            print(f"Failed to register security DLL in registry: {e}")
            print("The program will still work, but you may see a security popup once.")
    else:
        print("Warning: FilePathCheckerModuleExample.dll was not found on your system.")
        print("The program will still work, but you may need to click 'Always Allow' on the security popup when converting.")
        
    print("\nInstallation successful!")
    print("You can now right-click any HWP or HWPX file and click 'PDF로 변환' (under 'Show more options' in Windows 11).")
    return True

def unregister():
    extensions = [".hwp", ".hwpx"]
    
    # 1. Unregister Context Menu
    for ext in extensions:
        key_path = f"Software\\Classes\\SystemFileAssociations\\{ext}\\shell\\ConvertToPDF"
        
        # We need to delete subkeys first, then the key itself
        try:
            # Delete command subkey
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, f"{key_path}\\command")
            except FileNotFoundError:
                pass
            
            # Delete main shell key
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
                print(f"Unregistered context menu for {ext}")
            except FileNotFoundError:
                pass
        except Exception as e:
            print(f"Failed to unregister context menu for {ext}: {e}")
            
    # 2. Clean up security module entry from registry
    try:
        modules_path = r"SOFTWARE\HNC\HwpAutomation\Modules"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, modules_path, 0, winreg.KEY_SET_VALUE)
        try:
            winreg.DeleteValue(key, "FilePathCheckerModuleExample")
            print("Cleaned up Hancom Security Module DLL registry entry.")
        except FileNotFoundError:
            pass
        winreg.CloseKey(key)
    except Exception:
        # Key might not exist or already be deleted
        pass
        
    print("\nUninstallation successful!")
    return True

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="HWP to PDF Context Menu Setup")
    parser.add_argument("--uninstall", action="store_true", help="Uninstall the context menu and settings")
    args = parser.parse_args()
    
    if args.uninstall:
        unregister()
    else:
        register()
