import unittest

import app


class RequiredObjectiveTests(unittest.TestCase):
    def test_problem_only_proposes_objective_for_validation(self):
        result = app.digest_text(
            "Appointment delays and fragmented intake are limiting patient access.",
            "brief.txt",
        )
        self.assertTrue(result["case_objective"])
        self.assertEqual(result["objective_status"], "proposed")
        self.assertTrue(result["objective_requires_confirmation"])
        self.assertIn("Appointment delays", result["case_objective"])

    def test_unknown_goal_requires_input_not_invention(self):
        result = app.digest_text("Client: Example Organization", "brief.txt")
        self.assertEqual(result["case_objective"], "")
        self.assertEqual(result["objective_status"], "needs input")

    def test_formatted_heading_and_inline_scope_boundary(self):
        result = app.digest_text(
            "**Objectives**\nImprove patient access across outpatient services\n"
            "Scope: Scheduling systems and clinic operations",
            "brief.txt",
        )
        self.assertIn("Improve patient access", result["case_objective"])
        self.assertNotIn("Scheduling", result["case_objective"])

    def test_missing_objective_blocks_downstream_steps(self):
        client = app.app.test_client()
        for path in ("/api/draft", "/api/storyline", "/api/render"):
            for objective in ("", "   ", None):
                with self.subTest(path=path, objective=objective):
                    response = client.post(path, json={
                        "digest": {"case_objective": objective},
                    })
                    self.assertEqual(response.status_code, 422)
                    self.assertIn("objective", response.json["error"])

    def test_confirmed_objective_allows_draft(self):
        response = app.app.test_client().post("/api/draft", json={
            "digest": {"case_objective": "Improve patient access"},
        })
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json["case_objective"], "Improve patient access")