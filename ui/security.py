"""Security helpers for file handling, HTML output, and Excel sanitization."""
import html
import os
import zipfile
from typing import Optional


MAX_UPLOAD_SIZE = 50 * 1024 * 1024  # 50 MB
MAX_ZIP_FILES = 1000
MAX_ZIP_SIZE = 200 * 1024 * 1024  # 200 MB


class SecurityError(ValueError):
    """Raised when input validation or a security check fails."""
    pass


def sanitize_html(text) -> str:
    """Escape a value for safe insertion into HTML via unsafe_allow_html."""
    return html.escape(str(text), quote=True)


def sanitize_excel_value(value):
    """Prefix strings that could be interpreted as Excel formulas with a single quote."""
    if isinstance(value, str) and value:
        if value[0] in {"=", "+", "-", "@", "\n", "\r", "\t"}:
            return "'" + value
    return value


def safe_filename(name: str, default: str = "download") -> str:
    """Return a safe filename by removing path separators and null bytes."""
    if not isinstance(name, str):
        name = str(name)
    name = name.replace("\x00", "")
    name = name.replace("/", "_").replace("\\", "_")
    name = name.strip(" .")
    if not name:
        return default
    return name


def validate_upload_size(file_obj, max_size: int = MAX_UPLOAD_SIZE) -> None:
    """Validate that an uploaded file does not exceed the size limit."""
    if file_obj is None:
        return
    size = getattr(file_obj, "size", None)
    if size is not None and size > max_size:
        raise SecurityError(f"Uploaded file exceeds {max_size / 1e6:.0f} MB limit.")


def safe_extract_zip(zip_file_obj, dest_dir: str, *, max_files: int = MAX_ZIP_FILES, max_size: int = MAX_ZIP_SIZE) -> None:
    """Extract a ZIP archive safely, rejecting traversal, symlinks, and zip bombs."""
    real_dest = os.path.realpath(dest_dir)
    with zipfile.ZipFile(zip_file_obj) as zf:
        total_size = 0
        file_count = 0
        for info in zf.infolist():
            if info.is_dir():
                continue
            if info.file_size < 0 or info.compress_size < 0:
                raise SecurityError("Invalid ZIP entry metadata.")
            if ".." in info.filename or os.path.isabs(info.filename) or info.filename.startswith("/"):
                raise SecurityError(f"ZIP entry '{info.filename}' is not allowed (path traversal).")
            target = os.path.realpath(os.path.join(dest_dir, info.filename))
            if os.path.commonpath([target, real_dest]) != real_dest:
                raise SecurityError(f"ZIP entry '{info.filename}' resolves outside target directory.")
            total_size += info.file_size
            if total_size > max_size:
                raise SecurityError(f"ZIP contents exceed {max_size / 1e6:.0f} MB.")
            file_count += 1
            if file_count > max_files:
                raise SecurityError(f"ZIP contains more than {max_files} files.")
        zf.extractall(dest_dir)


def validate_data_dir(path: str) -> Optional[str]:
    """Validate that a user-supplied data directory is safe and not a system root."""
    if not path or not os.path.isdir(path):
        return f"Path '{path}' is not a valid directory."
    real_path = os.path.realpath(path)
    home_dir = os.path.realpath(os.path.expanduser("~"))
    blocked = {
        "/",
        "/etc",
        "/usr",
        "/var",
        "/bin",
        "/sbin",
        "/lib",
        "/lib64",
        "/opt",
        "/sys",
        "/proc",
        "/dev",
        home_dir,
    }
    if real_path in blocked:
        return f"Path '{path}' is not allowed as a data directory."
    return None
