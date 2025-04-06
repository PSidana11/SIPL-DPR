import subprocess

scripts = [
    "agra_finance.py",
    "agra.py",
    "nilothi_finance.py",
    "pkn_finance.py",   
    "pkn.py",
    "nilothi.py"
    "bilaspur.py",
    "bilaspur_finance.py",
    "dvc.py",
    "dvc_finance.py",
    "gpr.py",
    "lara.py",
    "lara_finance.py",
    "sitamarhi.py",
    "sitamarhi_finance.py",
    "yamunavihar.py",
    "yamunavihar_finance.py"
]

for script in scripts:
    print(f"Running {script}...")
    subprocess.run(["python", script], check=True)
    print(f"{script} completed.\n")
