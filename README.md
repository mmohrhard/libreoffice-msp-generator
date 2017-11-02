# libreoffice-msp-generator
Generator for LibreOffice MSP files from MSI files

## Necessary tools

* msimsp.exe (in PATH)
* msidb.exe (in PATH)
* msiexec.exe (in PATH)
* python3

## Documentation of the command line options

### -n / --new

Mandatory command line parameter with a mandatory additional parameter that points to the msi file that will be the updated version of the program.

### -o / --old

Mandatory command line parameter with a mandatory additional parameter that points to the msi file that represents the old version of the program.

### -l / --log

Enable the log output. Will print many additional logging messages.

### -s / --sign

Whether the generated MSP file should be signed.
