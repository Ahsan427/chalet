# chalet

Fill the Endpoint column on an NPI spreadsheet. Name, NPI, phone, address, and the other columns stay as they are.

The value is the **Endpoint** cell from the NPPES provider page (`table table-bordered table-sm`, Health Information Exchange). It is read from the public NPPES API, which returns that same field.

```bash
python3 -m pip install -r requirements.txt
python3 fill_endpoints.py clinics.csv -o Clinics_with_endpoints.xlsx --cache endpoints_cache.json
```

Several endpoints for one NPI are joined with `; `. A provider with no endpoint on the registry keeps a blank cell. A lookup that fails leaves that cell unchanged.
