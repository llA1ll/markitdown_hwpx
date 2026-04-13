import sys

from pathlib import Path

from ._exceptions import FileConversionException


def convert_hwp_to_hwpx(
    path: str | Path, target_path: str | Path | None = None
) -> Path:
    """
    Convert a .hwp document to a local .hwpx file.

    This conversion is supported only on Windows via Hancom Office COM automation.
    By default, the output is written next to the source file with the .hwpx suffix.
    """
    source_path = Path(path)
    target_path = Path(target_path) if target_path is not None else source_path.with_suffix(".hwpx")

    if sys.platform != "win32":
        raise FileConversionException(
            "Automatic .hwp conversion is only supported on Windows via Hancom COM. "
            "On this OS, please convert the file to .hwpx first and then run MarkItDown."
        )

    try:
        import pythoncom  # type: ignore
        import win32com.client as win32  # type: ignore
    except ImportError as exc:
        raise FileConversionException(
            "Automatic .hwp conversion requires pywin32 on Windows. "
            "Install pywin32, or convert the file to .hwpx before running MarkItDown."
        ) from exc

    if target_path.exists() and target_path.stat().st_size > 0:
        return target_path

    pythoncom.CoInitialize()
    hwp = None
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            # Some installations may not require or expose this registration.
            pass

        hwp.Open(str(source_path))
        hwp.SaveAs(str(target_path), "HWPX")
        hwp.Run("FileClose")
    except Exception as exc:
        raise FileConversionException(
            "Failed to convert .hwp to .hwpx using Hancom COM. "
            "Verify that Hancom Office is installed and COM automation is available."
        ) from exc
    finally:
        try:
            if hwp is not None:
                hwp.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    if not target_path.exists() or target_path.stat().st_size == 0:
        raise FileConversionException(
            "Hancom COM conversion finished, but no .hwpx output was produced."
        )

    return target_path
