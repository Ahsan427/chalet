import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook

from fill_endpoints import (
    endpoints_from_payload,
    fill_rows,
    find_columns,
    format_endpoints,
    normalize_npi,
    run,
)


class NormalizeNpiTest(unittest.TestCase):
    def test_text_and_excel_numbers(self):
        self.assertEqual(normalize_npi("1669475828"), "1669475828")
        self.assertEqual(normalize_npi(1669475828), "1669475828")
        self.assertEqual(normalize_npi(1669475828.0), "1669475828")
        self.assertEqual(normalize_npi("1.669475828E9"), "1669475828")
        self.assertEqual(normalize_npi("1669-475-828"), "1669475828")
        self.assertIsNone(normalize_npi(""))
        self.assertIsNone(normalize_npi(None))
        self.assertIsNone(normalize_npi("not an npi"))

    def test_columns(self):
        headers = ["Name", "NPI", "Phone number", "Location", "endpoints"]
        self.assertEqual(find_columns(headers), (1, 4))


class FillRowsTest(unittest.TestCase):
    def test_only_endpoint_column_changes(self):
        rows = [
            ["Name", "NPI", "Phone number", "Location", "endpoints"],
            ["CHERYL R CAMPBELL", "1669475828", "907-522-0078", "Anchorage, AK", None],
            ["Dr. JAMES L STRONG", "1619970290", "205-971-3600", "Birmingham, AL", ""],
            ["No number", "", "555-0100", "Somewhere", ""],
        ]
        original = [list(row) for row in rows]
        stats = fill_rows(
            rows,
            npi_index=1,
            endpoint_index=4,
            looked_up={
                "1669475828": [],
                "1619970290": ["james.strong.1@9954.direct.athenahealth.com"],
            },
            errors={},
            separator="; ",
        )
        self.assertEqual(rows[1][:4], original[1][:4])
        self.assertEqual(rows[2][:4], original[2][:4])
        self.assertEqual(rows[1][4], "")
        self.assertEqual(rows[2][4], "james.strong.1@9954.direct.athenahealth.com")
        self.assertEqual(rows[3], original[3])
        self.assertEqual(stats["filled"], 1)
        self.assertEqual(stats["blank"], 1)
        self.assertEqual(stats["skipped_no_npi"], 1)

    def test_multiple_endpoints_keep_registry_order(self):
        payload = {
            "results": [
                {
                    "number": "1234567893",
                    "endpoints": [
                        {"endpoint": "a@direct.example", "endpointType": "DIRECT"},
                        {"endpoint": "https://example.org/fhir", "endpointType": "FHIR"},
                        {"endpoint": "a@direct.example"},
                    ],
                }
            ]
        }
        values = endpoints_from_payload(payload, "1234567893")
        self.assertEqual(format_endpoints(values, "; "), "a@direct.example; https://example.org/fhir")

    def test_lookup_error_does_not_clear_the_cell(self):
        rows = [
            ["NPI", "Endpoint"],
            ["1234567893", "already here"],
        ]
        fill_rows(
            rows,
            npi_index=0,
            endpoint_index=1,
            looked_up={},
            errors={"1234567893": "timeout"},
            separator="; ",
        )
        self.assertEqual(rows[1][1], "already here")


class RunTest(unittest.TestCase):
    def test_run_writes_endpoint_and_leaves_other_columns(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / "clinics.csv"
            output = Path(directory) / "out.xlsx"
            source.write_text(
                "Name,NPI,Phone number,Phone extension,Location,Selected Taxonomy,endpoints\n"
                "Dr. JAMES L STRONG,1619970290,205-971-3600,,Birmingham,Hospitalist,\n"
                "CHERYL R CAMPBELL,1669475828,907-522-0078,,Anchorage,Speech-Language Pathologist,\n",
                encoding="utf-8",
            )

            def fake_fetch(npi: str, timeout: float = 30.0, attempts: int = 5):
                if npi == "1619970290":
                    return ["james.strong.1@9954.direct.athenahealth.com"]
                return []

            import fill_endpoints

            original = fill_endpoints.fetch_endpoints
            fill_endpoints.fetch_endpoints = fake_fetch
            try:
                stats = fill_endpoints.run(
                    source,
                    output,
                    sheet_name=None,
                    workers=1,
                    separator="; ",
                    cache_path=None,
                    pause=0,
                )
            finally:
                fill_endpoints.fetch_endpoints = original

            self.assertEqual(stats["filled"], 1)
            self.assertEqual(stats["blank"], 1)
            workbook = load_workbook(output)
            sheet = workbook.active
            headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
            self.assertEqual(
                headers,
                ["Name", "NPI", "Phone number", "Phone extension", "Location", "Selected Taxonomy", "endpoints"],
            )
            data = list(sheet.iter_rows(min_row=2, values_only=True))
            self.assertEqual(data[0][1], "1619970290")
            self.assertEqual(data[0][2], "205-971-3600")
            self.assertEqual(data[0][6], "james.strong.1@9954.direct.athenahealth.com")
            self.assertEqual(data[1][0], "CHERYL R CAMPBELL")
            self.assertIn(data[1][6], ("", None))
            self.assertEqual(sheet["B2"].number_format, "@")


if __name__ == "__main__":
    unittest.main()
