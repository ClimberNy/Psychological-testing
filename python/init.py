import subprocess
def install_packages():
    packages = ["python-docx","docx2pdf","PyPDF2","pdf2docx"]
    command = ["pip", "install"] + packages
    subprocess.check_call(command)


install_packages()