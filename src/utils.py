import subprocess


def ps_run(cmd):
    completed = subprocess.run(["powershell", "-Command", cmd], capture_output=True)
    return completed
