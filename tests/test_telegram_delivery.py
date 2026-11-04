"""Offline regression tests; never import main (it initializes the live DB)."""
import ast
import io
import json
import logging
from pathlib import Path
from types import SimpleNamespace
from typing import Dict, List, Optional
import unittest
from unittest.mock import Mock

import requests


SOURCE = Path(__file__).resolve().parents[1] / "main.py"
TOKEN = "123456:never-log-this-secret"


class Clock:
    def __init__(self):
        self.now = 0.0
        self.sleeps = []

    def monotonic(self):
        return self.now

    def sleep(self, seconds):
        self.sleeps.append(seconds)
        self.now += seconds


class Response:
    def __init__(self, payload, status=200):
        self.payload = payload
        self.status_code = status
        self.closed = False

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.closed = True

    def json(self):
        if isinstance(self.payload, Exception):
            raise self.payload
        return self.payload


def isolated_namespace():
    names = {
        "tg_call", "send_media", "tg_send_photo", "tg_send_video",
        "tg_send_animation", "tg_send_document", "send_scanner_result",
        "send_scanner_result_edit", "run_scanner_flow", "_coerce_media_dict",
    }
    constants = {
        "TG_API", "TG_HTTP_MAX_ATTEMPTS", "TG_HTTP_BUDGET_SECONDS",
        "TG_HTTP_RETRY_DELAY_SECONDS", "TG_RETRY_SAFE_METHODS",
    }
    nodes = [node for node in ast.parse(SOURCE.read_text()).body
             if (isinstance(node, ast.FunctionDef) and node.name in names)
             or (isinstance(node, ast.Assign) and any(
                 isinstance(target, ast.Name) and target.id in constants
                 for target in node.targets))]
    stream = io.StringIO()
    logger = logging.Logger("telegram-delivery-test")
    logger.addHandler(logging.StreamHandler(stream))
    ns = {
        "requests": SimpleNamespace(exceptions=requests.exceptions, post=Mock()),
        "SESSION": SimpleNamespace(post=Mock()), "time": Clock(),
        "logger": logger, "logs": stream, "json": json,
        "Dict": Dict, "List": List, "Optional": Optional,
        "TG_MAX_CAPTION": 1024,
        "sanitize_telegram_html": lambda value, **kwargs: value,
        "norm_provider": lambda value: value,
        "build_scanner_caption": Mock(return_value="result"),
        "build_scanner_result_keyboard": Mock(return_value={}),
        "get_scanner_duration_seconds": Mock(return_value=5),
        "animate_scanning_progress": Mock(),
    }
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(SOURCE), "exec"), ns)
    return ns


class TelegramRetryTests(unittest.TestCase):
    def setUp(self):
        self.ns = isolated_namespace()
        self.post = self.ns["SESSION"].post

    def call(self, method="sendMessage", **kwargs):
        return self.ns["tg_call"](TOKEN, method, **kwargs)

    def test_connect_timeout_recovers_once_without_duplicate_success(self):
        response = Response({"ok": True, "result": {"message_id": 7}})
        self.post.side_effect = [requests.exceptions.ConnectTimeout(TOKEN), response]
        self.assertEqual(self.call(), {"message_id": 7})
        self.assertEqual(self.post.call_count, 2)
        self.assertTrue(response.closed)
        self.assertEqual(self.ns["time"].sleeps, [0.25])

    def test_ambiguous_sends_and_payments_are_not_retried(self):
        for method in ("sendMessage", "sendPhoto", "sendVideo", "sendDocument", "sendInvoice"):
            for error in (requests.exceptions.ReadTimeout, requests.exceptions.ConnectionError):
                with self.subTest(method=method, error=error):
                    self.post.reset_mock(side_effect=True)
                    self.post.side_effect = error(TOKEN)
                    self.assertIsNone(self.call(method))
                    self.assertEqual(self.post.call_count, 1)

    def test_idempotent_edit_and_ack_recover_after_ambiguous_failure(self):
        for method in self.ns["TG_RETRY_SAFE_METHODS"]:
            with self.subTest(method=method):
                self.post.reset_mock(side_effect=True)
                response = Response({"ok": True, "result": True})
                self.post.side_effect = [requests.exceptions.ConnectionError(TOKEN), response]
                self.assertTrue(self.call(method))
                self.assertEqual(self.post.call_count, 2)
                self.assertTrue(response.closed)

    def test_retry_is_capped_at_two_attempts(self):
        self.post.side_effect = requests.exceptions.ReadTimeout(TOKEN)
        self.assertIsNone(self.call("editMessageText"))
        self.assertEqual(self.post.call_count, 2)
        self.assertEqual(self.ns["time"].sleeps, [0.25])

    def test_retry_uses_remaining_shared_budget(self):
        def first_attempt(*args, **kwargs):
            self.ns["time"].now += 20
            raise requests.exceptions.ReadTimeout(TOKEN)

        self.post.side_effect = first_attempt
        self.assertIsNone(self.call("editMessageText"))
        self.assertEqual(self.post.call_count, 2)
        second_timeout = self.post.call_args_list[1].kwargs["timeout"]
        self.assertLessEqual(sum(second_timeout), 4.75)

    def test_exhausted_budget_does_not_retry(self):
        def fail(*args, **kwargs):
            self.ns["time"].now += 25
            raise requests.exceptions.ConnectTimeout(TOKEN)

        self.post.side_effect = fail
        self.assertIsNone(self.call())
        self.assertEqual(self.post.call_count, 1)
        self.assertEqual(self.ns["time"].sleeps, [])

    def test_upload_and_stream_are_never_replayed(self):
        for body in ({"files": {"document": io.BytesIO(b"file")}}, {"data": io.BytesIO(b"body")}):
            with self.subTest(body=list(body)):
                self.post.reset_mock(side_effect=True)
                self.post.side_effect = requests.exceptions.ConnectTimeout(TOKEN)
                self.assertIsNone(self.call("sendDocument", **body))
                self.assertEqual(self.post.call_count, 1)

    def test_not_modified_is_confirmed_edit_success(self):
        response = Response({"ok": False, "error_code": 400, "description": "Bad Request: message is not modified"})
        self.post.side_effect = [requests.exceptions.ReadTimeout(TOKEN), response]
        self.assertTrue(self.call("editMessageText"))
        self.assertTrue(response.closed)
        self.assertEqual(self.post.call_count, 2)

    def test_429_remains_single_attempt_and_response_closed(self):
        response = Response({"ok": False, "error_code": 429, "parameters": {"retry_after": 34}})
        self.post.return_value = response
        self.assertIsNone(self.call("sendVideo"))
        self.assertEqual(self.post.call_count, 1)
        self.assertTrue(response.closed)
        self.assertEqual(self.ns["time"].sleeps, [])

    def test_existing_success_and_terminal_errors(self):
        for payload, expected in [
            ({"ok": True, "result": {"id": 7}}, {"id": 7}),
            ({"ok": False, "error_code": 403, "description": "bot was blocked"}, None),
            ({"ok": False, "error_code": 400, "description": "chat not found"}, None),
        ]:
            with self.subTest(payload=payload):
                self.post.reset_mock()
                response = Response(payload)
                self.post.return_value = response
                self.assertEqual(self.call(), expected)
                self.assertEqual(self.post.call_count, 1)
                self.assertTrue(response.closed)

    def test_logs_never_include_token_url_or_response_body(self):
        self.post.side_effect = requests.exceptions.ConnectionError("https://api.telegram.org/bot" + TOKEN)
        self.call()
        self.post.side_effect = None
        self.post.return_value = Response(ValueError(TOKEN), status=502)
        self.call()
        self.post.return_value = Response({"ok": False, "error_code": 400, "description": TOKEN})
        self.call()
        logs = self.ns["logs"].getvalue()
        self.assertIn("bot_id=123456", logs)
        self.assertIn("method=sendMessage", logs)
        self.assertIn("type=ConnectionError", logs)
        self.assertNotIn("never-log-this-secret", logs)
        self.assertNotIn("https://", logs)


class ScannerDeliveryTests(unittest.TestCase):
    def setUp(self):
        self.ns = isolated_namespace()
        self.events = []
        for name in ("edit_media", "edit_caption", "edit_message"):
            self.ns[name] = Mock(return_value=None)
        self.ns["delete_message"] = Mock(side_effect=lambda *args: self.events.append("delete"))
        self.ns["send_message"] = Mock(side_effect=lambda *args, **kwargs: self.events.append("send"))
        self.args = (TOKEN, 1, 2, "name", "provider", {}, [])

    def test_failed_replacement_retains_old_message_and_reports_failure(self):
        self.assertFalse(self.ns["send_scanner_result_edit"](*self.args))
        self.assertEqual(self.events, ["send"])
        self.ns["delete_message"].assert_not_called()

    def test_successful_replacement_is_sent_before_old_message_removed(self):
        def send(*args, **kwargs):
            self.events.append("send")
            return {"message_id": 3}

        self.ns["send_message"].side_effect = send
        self.assertTrue(self.ns["send_scanner_result_edit"](*self.args))
        self.assertEqual(self.events, ["send", "delete"])

    def test_confirmed_edit_does_not_send_or_delete(self):
        self.ns["edit_caption"].return_value = True
        self.assertTrue(self.ns["send_scanner_result_edit"](*self.args))
        self.assertEqual(self.events, [])

    def test_run_flow_does_not_send_second_fallback(self):
        self.ns["send_scanner_result_edit"] = Mock(return_value=False)
        self.ns["send_scanner_result"] = Mock()
        self.assertFalse(self.ns["run_scanner_flow"](*self.args))
        self.ns["send_scanner_result_edit"].assert_called_once()
        self.ns["send_scanner_result"].assert_not_called()

    def test_media_result_propagates_through_existing_wrappers(self):
        for media_type in ("photo", "video", "animation", "document"):
            for expected in (None, {"message_id": 3}):
                with self.subTest(media_type=media_type, expected=expected):
                    self.ns["tg_call"] = Mock(return_value=expected)
                    result = self.ns["send_scanner_result"](
                        TOKEN, 1, "name", "provider", {"media_type": media_type, "file_id": "file"}, [])
                    self.assertEqual(result, expected)
                    self.ns["tg_call"].assert_called_once()

    def test_failed_media_replacement_does_not_remove_old_message(self):
        self.ns["tg_call"] = Mock(return_value=None)
        args = (*self.args[:5], {"media_type": "photo", "file_id": "file"}, [])
        self.assertFalse(self.ns["send_scanner_result_edit"](*args))
        self.ns["tg_call"].assert_called_once()
        self.ns["delete_message"].assert_not_called()


if __name__ == "__main__":
    unittest.main()
