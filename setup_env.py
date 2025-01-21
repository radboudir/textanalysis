import os
import subprocess
import sys

def check_and_install_packages():
    # Install remaining packages from requirements.txt without printing "Requirement already satisfied"
    try:
        subprocess.check_call(
            [f"{venv_dir}/Scripts/python" if os.name == "nt" else f"{venv_dir}/bin/python",
             "-m", "pip", "install", "--quiet", "-r", "requirements.txt"]
        )
    except subprocess.CalledProcessError as e:
        sys.exit(f"Error installing packages: {e}")

def create_venv():
    global venv_dir
    venv_dir = "venv"

    if not os.path.exists(venv_dir):
        # Create virtual environment
        subprocess.check_call([sys.executable, "-m", "venv", venv_dir])
        print(f"Virtual environment created at {venv_dir}")
    else:
        print(f"Virtual environment already exists at {venv_dir}")

    check_and_install_packages()

def main():
    create_venv()
    
    # Print activation instructions
    if os.name == "nt":
        print(f"To activate the environment, run:\n{venv_dir}\\Scripts\\activate")
    else:
        print(f"To activate the environment, run:\nsource {venv_dir}/bin/activate")
    
    print("Virtual environment setup complete. Now you can activate it manually.")

if __name__ == "__main__":
    main()
