import subprocess

with open("requirements.txt") as f:
    lines = [line.strip() for line in f if line.strip() and not line.startswith("#")]

for pkg in lines:
    print(f"Installing {pkg} ...")
    try:
        subprocess.run(["pip", "install", pkg], check=True)
    except subprocess.CalledProcessError:
        print(f"Failed: {pkg}")
