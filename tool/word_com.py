from __future__ import annotations

import time
from typing import Any, Optional


def create_word_app(*, visible: bool = False) -> Any:
    """Create an isolated Word.Application with dialogs suppressed.

    Invisible modal dialogs (ConfirmConversions / encoding / alerts) are a common
    cause of COM calls hanging forever when ``Visible=False``.
    """
    import win32com.client

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = bool(visible)
    try:
        word.DisplayAlerts = 0  # wdAlertsNone
    except Exception:
        pass
    try:
        word.ScreenUpdating = bool(visible)
    except Exception:
        pass
    return word


def open_document(word: Any, path: str, *, read_only: bool = False) -> Any:
    """Open a document with flags that avoid interactive prompts."""
    return word.Documents.Open(
        FileName=str(path),
        ConfirmConversions=False,
        ReadOnly=bool(read_only),
        AddToRecentFiles=False,
        NoEncodingDialog=True,
    )


def release_word(
    word: Optional[Any] = None,
    doc: Optional[Any] = None,
    *,
    save_changes: bool = False,
    settle_seconds: float = 0.35,
) -> None:
    """Close document / quit Word and briefly wait so file locks release."""
    if doc is not None:
        try:
            if save_changes:
                try:
                    doc.Save()
                except Exception:
                    pass
            doc.Close(SaveChanges=bool(save_changes))
        except Exception:
            pass
    if word is not None:
        try:
            word.Quit()
        except Exception:
            pass
    if settle_seconds and settle_seconds > 0:
        time.sleep(float(settle_seconds))
