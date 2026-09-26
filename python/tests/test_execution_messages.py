"""Execution refusals should name the condition the caller can act on."""

from __future__ import annotations

from xlide_mcp.hosts import host_info
from xlide_mcp.tools.execution import _harness_refusal


def test_session_lock_refusal_does_not_claim_excel_is_open() -> None:
    info = host_info("Budget.xlsm")
    message = _harness_refusal(
        RuntimeError("Another pyvbaharness session is running for this application."), info
    )
    assert "execution lock" in message
    assert "Wait for it to finish and retry" in message
    assert "close Excel" not in message
