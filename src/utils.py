import subprocess


# Run a command in PowerShell
def ps_run(cmd: str) -> subprocess.CompletedProcess:
    completed = subprocess.run(["powershell", "-Command", cmd], capture_output=True)
    return completed


# Mount an ISO file in Windows
def mount_win(abspath: str) -> (str, str):
    cmd = ("$result = Mount-DiskImage \"{}\" -PassThru ;" +
           "$drive = ($result | Get-Volume) ;" +
           "").format(abspath)
    output = ps_run(cmd)
    if output.stderr != b'':
        print(output.stderr.decode("utf-8"))
        return False
    drive = output.stdout.decode("utf-8").strip()
    return drive


