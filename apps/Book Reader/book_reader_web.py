"""The Book Reader's private web set-up.

Book pages are shown by Qt WebEngine (Chromium). They come from a private
"book:" address scheme that serves the open book's chapters and pictures from
memory, through a profile that keeps nothing on disk. The book's own scripts
never run, and nothing in a book can reach the internet: only book:, qrc: and
data: requests load. The reader's own page script runs in a separate script
world (APP_WORLD), which keeps working with the book's scripts switched off.
"""

from __future__ import annotations

from typing import Callable

from PySide6.QtCore import QBuffer, QByteArray, QIODevice, QObject
from PySide6.QtWebEngineCore import (
    QWebEngineProfile,
    QWebEngineScript,
    QWebEngineSettings,
    QWebEngineUrlRequestInfo,
    QWebEngineUrlRequestInterceptor,
    QWebEngineUrlRequestJob,
    QWebEngineUrlScheme,
    QWebEngineUrlSchemeHandler,
)

SCHEME = b"book"
APP_WORLD = QWebEngineScript.ScriptWorldId.ApplicationWorld
ALLOWED_SCHEMES = {"book", "qrc", "data", "blob"}

Resolver = Callable[[str], "tuple[bytes, str] | None"]


def register_scheme() -> None:
    """Tell Qt about book: addresses. Must run before QApplication is made."""
    if QWebEngineUrlScheme.schemeByName(SCHEME).name() == SCHEME:
        return
    scheme = QWebEngineUrlScheme(SCHEME)
    scheme.setSyntax(QWebEngineUrlScheme.Syntax.Host)
    scheme.setFlags(
        QWebEngineUrlScheme.Flag.SecureScheme
        | QWebEngineUrlScheme.Flag.LocalScheme
        | QWebEngineUrlScheme.Flag.LocalAccessAllowed
        | QWebEngineUrlScheme.Flag.CorsEnabled
    )
    QWebEngineUrlScheme.registerScheme(scheme)


def request_path(url) -> str:
    """'book://b/text/one.html' -> 'b/text/one.html' (decoded, no query)."""
    from PySide6.QtCore import QUrl
    path = url.path(QUrl.ComponentFormattingOption.FullyDecoded).lstrip("/")
    return f"{url.host()}/{path}"


class BookSchemeHandler(QWebEngineUrlSchemeHandler):
    """Answers book: requests from the resolver callback."""

    def __init__(self, resolver: Resolver, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self.resolver = resolver

    def requestStarted(self, job: QWebEngineUrlRequestJob) -> None:  # noqa: N802 (Qt name)
        try:
            found = self.resolver(request_path(job.requestUrl()))
        except Exception:  # a damaged book must never take the window down
            found = None
        if found is None:
            job.fail(QWebEngineUrlRequestJob.Error.UrlNotFound)
            return
        data, mime = found
        buffer = QBuffer(job)
        buffer.setData(QByteArray(data))
        buffer.open(QIODevice.OpenModeFlag.ReadOnly)
        job.reply(mime.encode("ascii", "replace"), buffer)


class OfflineInterceptor(QWebEngineUrlRequestInterceptor):
    """Blocks every request that would leave the PC."""

    def interceptRequest(self, info: QWebEngineUrlRequestInfo) -> None:  # noqa: N802
        if info.requestUrl().scheme().lower() not in ALLOWED_SCHEMES:
            info.block(True)


def make_profile(parent: QObject, resolver: Resolver) -> QWebEngineProfile:
    """A profile that keeps nothing on disk, serves book: and stays offline."""
    profile = QWebEngineProfile(parent)  # no storage name: off the record
    profile.setHttpCacheType(QWebEngineProfile.HttpCacheType.MemoryHttpCache)
    profile.setPersistentCookiesPolicy(QWebEngineProfile.PersistentCookiesPolicy.NoPersistentCookies)
    handler = BookSchemeHandler(resolver, profile)
    profile.installUrlSchemeHandler(SCHEME, handler)
    interceptor = OfflineInterceptor(profile)
    profile.setUrlRequestInterceptor(interceptor)
    profile._book_handler = handler          # keep Python owners alive with the profile
    profile._book_interceptor = interceptor
    settings = profile.settings()
    attr = QWebEngineSettings.WebAttribute
    settings.setAttribute(attr.JavascriptEnabled, False)
    settings.setAttribute(attr.JavascriptCanOpenWindows, False)
    settings.setAttribute(attr.LocalContentCanAccessRemoteUrls, False)
    settings.setAttribute(attr.PluginsEnabled, False)
    settings.setAttribute(attr.AutoLoadIconsForPage, False)
    settings.setAttribute(attr.ErrorPageEnabled, False)
    settings.setAttribute(attr.FocusOnNavigationEnabled, True)
    return profile
