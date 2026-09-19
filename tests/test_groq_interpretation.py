import json
import sys
import types
import unittest
from copy import deepcopy
from unittest.mock import patch

import app


SOURCE_TEXT = """
Northstar Foods is a consumer packaged foods company seeking support for its India retail business.
The objective is to increase modern trade revenue by 15% within 12 months.
The work includes customer research and a channel growth roadmap.
"""


def groq_module_with_response(payload=None, error=None):
    class Completions:
        def create(self, **kwargs):
            if error:
                raise error
            message = types.SimpleNamespace(content=json.dumps(payload))
            return types.SimpleNamespace(
                choices=[types.SimpleNamespace(message=message)]
            )

    class Groq:
        def __init__(self, **kwargs):
            self.chat = types.SimpleNamespace(
                completions=Completions()
            )

    return types.SimpleNamespace(Groq=Groq)


class GroqInterpretationTests(unittest.TestCase):
    def setUp(self):
        self.baseline = app.digest_text(SOURCE_TEXT, "brief.txt")

    def interpret(self, payload=None, error=None, baseline=None):
        fake_groq = groq_module_with_response(payload, error)
        with (
            patch.dict("os.environ", {"GROQ_API_KEY": "test-key"}, clear=False),
            patch.dict(sys.modules, {"groq": fake_groq}),
        ):
            return app.interpret_digest_with_groq(
                SOURCE_TEXT,
                "",
                "brief.txt",
                deepcopy(baseline or self.baseline),
            )

    def test_grounded_model_response_overrides_deterministic_candidate(self):
        payload = {
            "client": "Northstar Foods",
            "sector": "Consumer packaged foods",
            "quantified_target": "Increase modern trade revenue by 15% within 12 months",
            "target_type": "quantified",
            "field_evidence": {
                "client": ["Northstar Foods"],
                "sector": ["consumer packaged foods"],
                "quantified_target": [
                    "The objective is to increase modern trade revenue by 15% within 12 months."
                ],
            },
        }

        result = self.interpret(payload)

        self.assertEqual(result["client"], "Northstar Foods")
        self.assertEqual(result["sector"], "Consumer packaged foods")
        self.assertEqual(
            result["quantified_target"],
            "Increase modern trade revenue by 15% within 12 months",
        )
        self.assertEqual(
            result["interpretation_source"],
            "Groq structured interpretation with deterministic fallback",
        )

    def test_unsupported_model_facts_retain_deterministic_values(self):
        baseline = deepcopy(self.baseline)
        payload = {
            "client": "Northstar Foods International",
            "sector": "Aerospace",
            "quantified_target": "Save $900 million",
            "target_type": "mixed",
            "scope_in": [
                "Customer research",
                "Customer research and nuclear reactor design",
            ],
            "field_evidence": {
                "client": ["Northstar Foods"],
                "sector": ["consumer packaged foods"],
                "quantified_target": [
                    "The objective is to increase modern trade revenue by 15% within 12 months."
                ],
                "scope_in": [
                    "The work includes customer research and a channel growth roadmap."
                ],
            },
        }

        result = self.interpret(payload, baseline=baseline)

        self.assertEqual(result["client"], baseline["client"])
        self.assertEqual(result["sector"], baseline["sector"])
        self.assertEqual(
            result["quantified_target"],
            baseline["quantified_target"],
        )
        self.assertEqual(result["target_type"], baseline["target_type"])
        self.assertEqual(
            result["scope_in"],
            ["customer research and a channel growth roadmap"],
        )

    def test_missing_configuration_and_api_failure_return_usable_fallback(self):
        baseline = deepcopy(self.baseline)
        with patch.dict("os.environ", {}, clear=True):
            unavailable = app.interpret_digest_with_groq(
                SOURCE_TEXT, "", "brief.txt", deepcopy(baseline)
            )

        failed = self.interpret(
            error=RuntimeError("model unavailable"),
            baseline=baseline,
        )

        for result in (unavailable, failed):
            self.assertEqual(result["client"], baseline["client"])
            self.assertEqual(
                result["quantified_target"],
                baseline["quantified_target"],
            )
            self.assertEqual(
                result["interpretation_source"],
                "deterministic fallback",
            )
            self.assertTrue(result["interpretation_warning"])

    def test_descriptive_percentage_is_not_promoted_to_target(self):
        text = (
            "Northstar Foods operates in India. Modern trade currently "
            "accounts for 15% of revenue. The scope includes customer research."
        )

        result = app.digest_text(text, "brief.txt")

        self.assertEqual(result["quantified_target"], "")
        self.assertEqual(result["quantified_target_evidence"], [])
        self.assertEqual(result["target_type"], "not identified")

    def test_qualitative_objective_has_its_own_field(self):
        text = (
            "The objective is to improve patient access and reduce appointment friction. "
            "The work includes scheduling processes and clinic operations."
        )

        result = app.digest_text(text, "healthcare-brief.txt")

        self.assertEqual(
            result["case_objective"],
            "The objective is to improve patient access and reduce appointment friction.",
        )
        self.assertEqual(result["quantified_target"], "")
        self.assertEqual(result["target_type"], "qualitative objective")

    def test_objective_heading_captures_following_bullets(self):
        text = """
        Background
        Appointment delays are limiting access.

        Objectives
        • Improve patient access across outpatient services
        • Reduce friction in the booking journey

        Scope
        Scheduling processes and clinic operations
        """

        result = app.digest_text(text, "healthcare-brief.txt")

        self.assertEqual(
            result["case_objective"],
            (
                "Improve patient access across outpatient services "
                "Reduce friction in the booking journey"
            ),
        )

    def test_inline_the_ask_is_parsed_as_an_objective(self):
        result = app.digest_text(
            "The Ask: Create a scalable operating model for regional service delivery.",
            "operating-model-brief.txt",
        )

        self.assertEqual(
            result["case_objective"],
            "Create a scalable operating model for regional service delivery.",
        )

    def test_scope_extraction_is_not_limited_to_fmcg_terms(self):
        text = (
            "Scope includes identity management systems, onboarding processes, "
            "and customer-support operations."
        )

        result = app.digest_text(text, "technology-brief.txt")

        self.assertEqual(
            result["scope_in"],
            [
                "identity management systems, onboarding processes, "
                "and customer-support operations"
            ],
        )


if __name__ == "__main__":
    unittest.main()