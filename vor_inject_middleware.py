"""
vor_inject_middleware.py — инжектит <script src="/static/vor_patch.js">
в HTML-ответы корня "/", не модифицируя сам шаблон index.html.

Подключается из run_web.py:
    from vor_inject_middleware import install_html_inject
    install_html_inject(app)
"""

from __future__ import annotations

from starlette.middleware.base import BaseHTTPMiddleware


_INJECT_TAG = '<script src="/static/vor_patch.js" defer></script>'


class _HtmlInjectMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        response = await call_next(request)
        # Инжектим только в HTML-ответы корня и индексной страницы.
        if request.url.path not in ("/", "/index", "/index.html"):
            return response

        ctype = response.headers.get("content-type", "")
        if "text/html" not in ctype:
            return response

        # Собираем тело
        body = b""
        async for chunk in response.body_iterator:
            body += chunk

        try:
            text = body.decode("utf-8")
        except UnicodeDecodeError:
            return _passthrough(response, body)

        if _INJECT_TAG in text:
            return _passthrough(response, body)

        # Вставляем перед </body> (или просто в конец, если тега нет).
        if "</body>" in text:
            text = text.replace("</body>", _INJECT_TAG + "\n</body>", 1)
        else:
            text += "\n" + _INJECT_TAG

        new_body = text.encode("utf-8")
        from starlette.responses import Response as _Resp
        # пересобираем заголовки без content-length (он изменится)
        headers = {k: v for k, v in response.headers.items()
                   if k.lower() not in ("content-length",)}
        return _Resp(
            content=new_body,
            status_code=response.status_code,
            headers=headers,
            media_type=ctype,
        )


def _passthrough(response, body: bytes):
    from starlette.responses import Response as _Resp
    headers = {k: v for k, v in response.headers.items()
               if k.lower() not in ("content-length",)}
    return _Resp(
        content=body,
        status_code=response.status_code,
        headers=headers,
        media_type=response.headers.get("content-type"),
    )


def install_html_inject(app):
    """Подключить инжектор к FastAPI-приложению."""
    app.add_middleware(_HtmlInjectMiddleware)
    return app
