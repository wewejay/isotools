import subprocess


# Run a command in PowerShell
def ps_run(cmd: str) -> subprocess.CompletedProcess:
    completed = subprocess.run(["powershell", "-Command", cmd], capture_output=True)
    return completed


# Mount an ISO file
def mount(abspath):
    cmd = "$result = Mount-DiskImage \"{}\" -PassThru ; ($mountResult | Get-Volume).DriveLetter".format(abspath)
    output = ps_run(cmd)
    if output.stderr != b'':
        print(output.stderr.decode("utf-8"))
        return False
    drive = output.stdout.decode("utf-8").strip()
    return drive

