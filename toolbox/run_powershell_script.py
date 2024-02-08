
import subprocess

def run_powershell_script(script_path):
    command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path]
    subprocess.run(command)
