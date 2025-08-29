# CSAM Inventory

## Installation of Dependencies

``` powershell
poetry install
```

### SSL/TLS Certificates

If installing dependencies while connect to the ED VPN, the 
`REQUESTS_CA_BUNDLE` environment variable must be set to point 
to a file containing the chain of certificates used by the firewall
in place of the actual certificates for pypi.org. To do this in
PowerShell, run the following

``` powershell
$env:REQUESTS_CA_BUNDLE="ext/certificates/pypi.cer"
```

Similar commands can be used with the Windows Command Prompt or 
other shells.

## Configuration

Configuration should be specified in a file named `config.yml` using YAML to
format the data; see `config.yml.example` for a sample configuration.

## Execution

To start the program, run

``` powershell
python inventory.py
```

### Optional Arguments

```text
  -h, --help       show the help message and exit
  --skip-download  skip downloading new hardware inventory files
  --config CONFIG  path to configuration file, default is ./config.yml
```

### PDF Conversion

In order to process data from PDFs, each PDF is converted to a Word document.
While this conversion is mostly automated, a user might see a dialog box similar
to the following:

![Word PDF Conversion Dialog](docs/images/word-pdf.png)

The user should select the ***Don't show this message again*** option and click
***OK***. This will prevent additional dialogs from appearing for additional PDF
conversions.

## Development

During development, a linter can be useful. To run pylint against the project
code, execute the following:

```powershell
pylint .\inventory.py
pylint .\csam_inventory\
```
