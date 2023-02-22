# https://stackoverflow.com/questions/52578270/install-python-with-cmd-or-powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Invoke-WebRequest -Uri "https://www.python.org/ftp/python/3.10.8/python-3.10.8.exe" -OutFile "c:/temp/python-3.10.8.exe"

c:/temp/python-3.10.8.exe /quiet InstallAllUsers=0 PrependPath=1 Include_test=0

# add to path
#$env:Path += <need to get path to installed python.exe for Python310 here>

# set up virtual environment:
python -m venv venv
# activate it:
./venv/Scripts/Activate

# install requirements
python -m pip install -r requirements.txt

write-host("All done!")
write-host("......................................................")
write-host("Try running loans.py with the command: python loans.py")
