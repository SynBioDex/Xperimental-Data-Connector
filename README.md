# Xperimental-data-convertor

This is a utility to link excel2sbol and excel to flapjack converter, interlink the resulting data, and upload it to synbiohub and flapjack.

## Plate-reader exports

Create a Flapjack-compatible workbook from an XDC template and a BioTek-style
CSV/TXT reader export. The template's `SBH_sampledesigns_collection` worksheet
must map its sample-design names to SynBioHub SBOL object URIs.

```python
import os

from xperimental_data_conv import create_flapjack_input

create_flapjack_input(
    "Align_TF_Study.xlsm",
    "2025-07-02_Align-TF-Cytom-12-plasmids_growth-plate_1-1_Neo.txt",
    "Flapjack_input.xlsx",
    username=os.environ["SBH_USERNAME"],
    password=os.environ["SBH_PASSWORD"],
)
```

The output has `Data`, `Media`, `Strains`, `DNA`, and `Chemicals` worksheets.
The `Data` worksheet retains the reader export; the other sheets are 96-well
maps resolved from the SBOL graph. Media uses role `NCIT:C48164`, strain
assemblies use role `NCIT:C14419`, chemicals use their functional-component
definition and `Flapjack#concentration`, and plasmids become `DNA 1...N` maps.
Use `token=os.environ["SBH_TOKEN"]` instead of username/password for token
authentication.
