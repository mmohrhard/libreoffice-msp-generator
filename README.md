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

## The following environment variables allow to control the MSP generation

* LIBO_SERVICEPACK: Default value: 'Hotfix'; Valid values: 'Hotfix', 'ServicePack'
* LIBO_ALLOWREMOVAL: Default value: '1'; Valid values: '0', '1'
* LIBO_PRODUCTNAME: Default value 'LibreOffice'
* LIBO_VENDOR: Default value 'LibreOffice'
* LIBO_PATCHLEVEL: Default value '0'; Valid values are non-negative integers
* LIBO_PACKAGEVERSION: Mandatory parameter with the package version, form $major.$minor.$micro.$patch
