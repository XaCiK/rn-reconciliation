# -*- coding: utf-8 -*-
"""Restore modReconcile.bas from subagent transcript with CP1251 encoding."""
import json
import re
from pathlib import Path

TRANSCRIPT = Path(
    r"C:\Users\o_antonov\.cursor\projects\c-Users-o-antonov-NEW-RN"
    r"\agent-transcripts\2228858f-2fa8-414c-a236-a9626eab9148"
    r"\subagents\a01a69d8-a862-4085-a93a-9c8cb0723cbf.jsonl"
)
OUT = Path(r"C:\Users\o_antonov\NEW\RN\modReconcile.bas")

FREEZE_OLD_BANK = """    wsRpt.Columns("A:P").AutoFit
    wsRpt.Rows("1:1").RowHeight = 28
    wsRpt.Range("A1").Select
    ActiveWindow.FreezePanes = False
    wsRpt.Range("A2").Select
    ActiveWindow.FreezePanes = True
    wsRpt.AutoFilterMode = False"""

FREEZE_NEW_BANK = """    wsRpt.Columns("A:P").AutoFit
    wsRpt.Rows("1:1").RowHeight = 28
    ApplyFreezeTopRow wsRpt
    wsRpt.AutoFilterMode = False"""

FREEZE_OLD_RN = """    wsRpt.Columns("A:M").AutoFit
    wsRpt.Columns("E").ColumnWidth = 35
    wsRpt.Rows("1:1").RowHeight = 28
    wsRpt.Range("A2").Select
    ActiveWindow.FreezePanes = True
    wsRpt.AutoFilterMode = False"""

FREEZE_NEW_RN = """    wsRpt.Columns("A:M").AutoFit
    wsRpt.Columns("E").ColumnWidth = 35
    wsRpt.Rows("1:1").RowHeight = 28
    ApplyFreezeTopRow wsRpt
    wsRpt.AutoFilterMode = False"""

APPLY_FREEZE = """
Private Sub ApplyFreezeTopRow(ws As Worksheet)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 1
    ActiveWindow.SplitColumn = 0
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

"""


def extract_contents() -> str:
    with TRANSCRIPT.open(encoding="utf-8") as f:
        for line in f:
            if '"name": "Write"' not in line or "modReconcile.bas" not in line:
                continue
            if "PopulateDefaultDictionary" not in line:
                continue
            data = json.loads(line)
            for block in data["message"]["content"]:
                if block.get("type") == "tool_use" and block.get("name") == "Write":
                    return block["input"]["contents"]
    raise RuntimeError("Could not find Write block in transcript")


def main() -> None:
    content = extract_contents()
    content = content.replace(FREEZE_OLD_BANK, FREEZE_NEW_BANK)
    content = content.replace(FREEZE_OLD_RN, FREEZE_NEW_RN)
    if "ApplyFreezeTopRow" not in content:
        content = content.replace(
            "Private Sub ColorBankRow(ws As Worksheet, rowNum As Long, status As String, isGSM As Boolean)",
            APPLY_FREEZE.strip() + "\n\nPrivate Sub ColorBankRow(ws As Worksheet, rowNum As Long, status As String, isGSM As Boolean)",
        )
    # Normalize line endings for VBA
    content = content.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    OUT.write_bytes(content.encode("cp1251"))
    text = OUT.read_text(encoding="cp1251")
    checks = ["Словарь", "нефтепродукт", "Зачислено", "ГСМ_РНКарт", "ApplyFreezeTopRow", "RN_Slovar"]
    for s in checks:
        ok = s in text
        print(f"{'OK' if ok else 'FAIL'}: {s}")
    print(f"Lines: {len(text.splitlines())}")
    print(f"Bytes: {OUT.stat().st_size}")


if __name__ == "__main__":
    main()
