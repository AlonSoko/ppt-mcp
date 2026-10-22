"""Downloads the tools make: icons, the icon list, pictures by URL.

Each has a time limit, so a network that drops the request rather than
refusing it answers in seconds instead of holding the call until the
operating system gives up. A download that cannot get through says which
host it tried and, for the icon set, which setting points it elsewhere.
"""

import os
import socket
import urllib.error
import urllib.parse
import urllib.request

# Seconds to wait for a server to answer. PPT_DOWNLOAD_TIMEOUT changes it.
DOWNLOAD_TIMEOUT = float(os.getenv("PPT_DOWNLOAD_TIMEOUT", "15"))


def open_url(url, what, setting=None, method=None, timeout=None):
    """``urllib.request.urlopen`` with a time limit and an error that helps.

    An HTTP error, a 404 above all, is passed through as it is, since the
    caller knows what a missing file means. Anything that kept the request
    from getting an answer at all becomes a ConnectionError naming the host,
    and the setting that moves the download to a mirror when there is one.
    A ``file://`` address that names no file is a 404, so a local copy of
    the icon set behaves like the real one.
    """
    request = urllib.request.Request(url, method=method) if method else url
    try:
        return urllib.request.urlopen(request, timeout=timeout or DOWNLOAD_TIMEOUT)
    except urllib.error.HTTPError:
        raise
    except urllib.error.URLError as exc:
        if isinstance(exc.reason, FileNotFoundError):
            raise urllib.error.HTTPError(url, 404, "No such file", None, None) from None
        raise ConnectionError(_unreachable(url, what, setting, exc.reason)) from None
    except (socket.timeout, TimeoutError, OSError) as exc:
        raise ConnectionError(_unreachable(url, what, setting, exc)) from None


def _unreachable(url, what, setting, reason):
    host = urllib.parse.urlsplit(url).netloc or url
    if isinstance(reason, (socket.timeout, TimeoutError)):
        reason = f"no answer within {DOWNLOAD_TIMEOUT:g} s"
    message = f"Could not reach {host} to {what} ({reason})."
    if setting:
        message += (
            f" If this network blocks {host}, set {setting} to a mirror or a "
            "local copy; the README says how."
        )
    return message
