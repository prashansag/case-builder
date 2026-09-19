import unittest
from unittest.mock import patch, call

import app


class ContextSourcesTests(unittest.TestCase):
    def test_only_supplied_urls_are_fetched_and_duplicates_removed(self):
        urls = ["https://client.example.com", "https://research.example.com/report"]
        def profile(url):
            return {"url": url, "title": "Source", "excerpt": "Evidence"}, None

        with patch("app.page_profile", side_effect=profile) as fetch, patch(
            "app.requests.get", side_effect=AssertionError("Unexpected network search")
        ):
            response = app.app.test_client().post("/api/context", json={
                "brand_url": urls[0], "urls": [urls[1], " " + urls[0] + " "],
                "digest": {"client": "Example", "sector": "Healthcare"},
            })
        self.assertEqual(response.status_code, 200)
        self.assertEqual(fetch.call_args_list, [call(url) for url in urls])
        self.assertEqual(len(response.json["sources"]), 2)

    def test_no_urls_means_no_research(self):
        with patch("app.requests.get", side_effect=AssertionError("Unexpected network")), patch(
            "app.page_profile", side_effect=AssertionError("Unexpected fetch")
        ):
            response = app.app.test_client().post("/api/context", json={
                "digest": {"client": "Example"}
            })
        self.assertEqual(response.json["sources"], [])

    def test_failed_supplied_url_is_not_replaced_by_search(self):
        with patch("app.page_profile", side_effect=ValueError("Unavailable")), patch(
            "app.requests.get", side_effect=AssertionError("Unexpected network search")
        ):
            response = app.app.test_client().post("/api/context", json={
                "urls": ["https://research.example.com/report"],
                "digest": {"client": "Example"},
            })
        self.assertEqual(len(response.json["sources"]), 1)
        self.assertEqual(response.json["sources"][0]["status"], "failed")