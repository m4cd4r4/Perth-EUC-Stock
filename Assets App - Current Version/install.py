import os
import subprocess
import sys
import shutil

def install_requirements():
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
    except subprocess.CalledProcessError as e:
        print(f"Failed to install requirements: {e}")
        sys.exit(1)

def copy_files():
    files_to_copy = [
        "icon.png",
        "build_roomv3.6.py",
        "inventory-levels_4.2v2.py",
        "inventory-levels_BRv2.py",
        "inventory-levels_combinedv1.2.py",
    ]
    for file in files_to_copy:
        if os.path.exists(file):
            shutil.copy(file, "venv/")

def create_shortcut():
    desktop = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')
    shortcut_path = os.path.join(desktop, 'Perth EUC Stock.lnk' if os.name == 'nt' else 'Perth EUC Stock.desktop')

    if os.name == 'nt':  # Windows
        import winshell
        from win32com.client import Dispatch

        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = os.path.join(os.getcwd(), 'venv\\Scripts\\python.exe')
        shortcut.Arguments = os.path.join(os.getcwd(), 'venv\\build_roomv3.6.py')
        shortcut.WorkingDirectory = os.path.join(os.getcwd(), 'venv')
        shortcut.IconLocation = os.path.join(os.getcwd(), 'venv\\icon.png')
        shortcut.save()
    else:  # Unix-based
        with open(shortcut_path, 'w') as shortcut:
            shortcut.write(f"""[Desktop Entry]
Type=Application
Name=Perth EUC Stock
Exec={os.path.join(os.getcwd(), 'venv/bin/python')} {os.path.join(os.getcwd(), 'venv/build_roomv3.6.py')}
Icon={os.path.join(os.getcwd(), 'venv/icon.png')}
Terminal=false
""")
        os.chmod(shortcut_path, 0o755)

def main():
    try:
        # Create a virtual environment
        subprocess.check_call([sys.executable, "-m", "venv", "venv"])
    except subprocess.CalledProcessError as e:
        print(f"Failed to create virtual environment: {e}")
        sys.exit(1)
    
    # Activate the virtual environment
    activate_script = os.path.join('venv', 'Scripts', 'activate') if os.name == 'nt' else os.path.join('venv', 'bin', 'activate')
    activate_command = f"source {activate_script}" if os.name != 'nt' else activate_script

    try:
        subprocess.check_call(activate_command, shell=True)
    except subprocess.CalledProcessError as e:
        print(f"Failed to activate virtual environment: {e}")
        sys.exit(1)

    # Install requirements
    install_requirements()
    
    # Copy necessary files
    copy_files()

    # Create desktop shortcut
    create_shortcut()

    print("Installation complete. A shortcut has been created on your desktop to run the application.")

if __name__ == "__main__":
    main()
