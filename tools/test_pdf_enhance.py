import importlib.util
import unittest
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "pdf_enhance.py"
SPEC = importlib.util.spec_from_file_location("pdf_enhance", SCRIPT)
assert SPEC and SPEC.loader
pdf_enhance = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(pdf_enhance)


class OcrCandidateScoreTests(unittest.TestCase):
    def test_prefers_candidate_with_more_distinct_table_values(self):
        incomplete = "Metric A\nMetric B 30,000"
        complete = "Metric A 49,000 30,000\nMetric B 20 10"

        self.assertGreater(
            pdf_enhance._ocr_candidate_score(complete),
            pdf_enhance._ocr_candidate_score(incomplete),
        )

    def test_distinct_values_beat_repeated_row_numbers(self):
        repeated_serials = "1 Metric A 1 0\n2 Metric B 2\nTotal 4,629"
        recovered_values = "Metric A 1 0\nMetric B 13 2\nTotal 4,629"

        self.assertGreater(
            pdf_enhance._ocr_candidate_score(recovered_values),
            pdf_enhance._ocr_candidate_score(repeated_serials),
        )


if __name__ == "__main__":
    unittest.main()
