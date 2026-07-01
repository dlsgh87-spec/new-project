from __future__ import annotations

import json
import time
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlparse


@dataclass
class LoginSignals:
    current_url: str
    is_login_page: bool
    cookie_count: int
    matched_signals: list[str]

    @property
    def looks_logged_in(self) -> bool:
        return bool(self.matched_signals)


def capture_storage_state(url: str, state_path: Path, headless: bool, label: str) -> None:
    from playwright.sync_api import Error, sync_playwright

    state_path.parent.mkdir(parents=True, exist_ok=True)
    session_storage_path = state_path.with_name(state_path.stem + "_session_storage.json")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=headless)
        context_kwargs = {}
        if state_path.exists():
            context_kwargs["storage_state"] = str(state_path)
        context = browser.new_context(**context_kwargs)

        action = {"kind": None}

        def request_save(_source: object, kind: str = "force") -> None:
            action["kind"] = kind

        login_page = context.new_page()
        login_page.goto(url, wait_until="domcontentloaded")

        control_page = context.new_page()
        control_page.expose_binding("requestSave", request_save)
        control_page.set_content(_control_html(label, state_path))
        _set_status(control_page, "1. 로그인 탭 열림")
        _append_status(control_page, "2. 사용자 로그인 대기 중")
        login_page.bring_to_front()

        try:
            completed = False
            last_signal_text = ""
            while browser.is_connected() and not completed:
                while action["kind"] is None:
                    time.sleep(0.75)
                    if not browser.is_connected():
                        raise RuntimeError(f"{label} 브라우저가 세션 저장 전에 닫혔습니다.")
                    try:
                        signals = _detect_login_signals(login_page, context, url)
                        signal_text = _signal_summary(signals)
                        if signal_text != last_signal_text:
                            _show_diagnostics(control_page, signals, state_path, "")
                            last_signal_text = signal_text
                    except Exception as exc:  # UI diagnostics should never stop the flow.
                        _append_status(control_page, f"로그인 상태 확인 중 오류: {exc}")

                control_page.bring_to_front()
                _append_status(control_page, "3. 세션 저장 시도 중")

                try:
                    signals = _detect_login_signals(login_page, context, url)
                    if state_path.exists():
                        state_path.unlink()
                    if session_storage_path.exists():
                        session_storage_path.unlink()
                    context.storage_state(path=str(state_path))
                    _save_session_storage(context.pages, session_storage_path)
                    _append_status(control_page, "4. storage state 파일 생성 확인 중")

                    if not state_path.exists() or state_path.stat().st_size == 0:
                        raise RuntimeError("storage state 파일이 생성되지 않았거나 비어 있습니다.")

                    _append_status(control_page, "5. 저장된 세션으로 재접속 테스트 중")
                    reconnect = _test_saved_session(browser, state_path, url)
                    _append_status(control_page, "6. 완료")

                    if reconnect.looks_logged_in:
                        _show_success(
                            control_page,
                            f"{label} 로그인 세션이 저장되었습니다.\n이제 재고 수집 테스트를 진행할 수 있습니다.\n\n세션 저장 성공 / 재접속 테스트 성공",
                            signals,
                            state_path,
                            session_storage_path,
                        )
                    else:
                        _show_success(
                            control_page,
                            f"{label} 로그인 세션이 저장되었습니다.\n이제 재고 수집 테스트를 진행할 수 있습니다.\n\n세션 저장은 되었지만 재접속 테스트 실패: 다시 로그인 필요",
                            reconnect,
                            state_path,
                            session_storage_path,
                        )
                    completed = True
                    time.sleep(4)
                except Exception as exc:
                    try:
                        signals = _detect_login_signals(login_page, context, url)
                    except Exception:
                        signals = LoginSignals(current_url=login_page.url, is_login_page=True, cookie_count=0, matched_signals=[])
                    _show_failure(control_page, signals, state_path, str(exc))
                    _append_status(control_page, "저장 실패. 사용자 재시도 대기 중")
                    action["kind"] = None
                    login_page.bring_to_front()
        finally:
            try:
                browser.close()
            except Error:
                pass


def _detect_login_signals(page: object, context: object, login_url: str) -> LoginSignals:
    from playwright.sync_api import Error

    current_url = getattr(page, "url", "")
    login_host = urlparse(login_url).netloc.lower()
    current = urlparse(current_url)
    current_host = current.netloc.lower()
    current_path = current.path.lower()
    is_login_page = "login" in current_path or current_url.rstrip("/") == login_url.rstrip("/")

    cookies = context.cookies()
    cookie_count = len(cookies)
    matched: list[str] = []

    if current_host and current_host == login_host and not is_login_page:
        matched.append("현재 URL이 로그인 페이지가 아님")
    if cookie_count > 0:
        matched.append("세션 쿠키가 존재함")

    selectors = [
        ("좌측 메뉴 또는 상단 메뉴가 보임", "nav, aside, [role='navigation'], .gnb, .lnb, .menu, #menu, #gnb, #lnb"),
        ("이카운트 메인 화면으로 진입함", "text=/ECOUNT|이카운트|ERP/i"),
        ("재고I 또는 재고현황 메뉴 접근 가능", "text=/재고I|재고현황|재고/i"),
        ("사용자/회사명이 표시됨", "text=/회사|사용자|로그아웃|Logout|My Page|마이페이지/i"),
    ]

    for message, selector in selectors:
        try:
            if page.locator(selector).first.count() > 0:
                matched.append(message)
        except Error:
            continue

    return LoginSignals(
        current_url=current_url,
        is_login_page=is_login_page,
        cookie_count=cookie_count,
        matched_signals=list(dict.fromkeys(matched)),
    )


def _test_saved_session(browser: object, state_path: Path, url: str) -> LoginSignals:
    from playwright.sync_api import TimeoutError

    test_context = browser.new_context(storage_state=str(state_path))
    try:
        page = test_context.new_page()
        page.goto(url, wait_until="domcontentloaded", timeout=30000)
        try:
            page.wait_for_load_state("networkidle", timeout=5000)
        except TimeoutError:
            pass
        return _detect_login_signals(page, test_context, url)
    finally:
        test_context.close()


def _save_session_storage(pages: list[object], path: Path) -> None:
    snapshots = []
    for page in pages:
        url = getattr(page, "url", "")
        if not url.startswith("http"):
            continue
        try:
            data = page.evaluate(
                """
                () => {
                  const values = {};
                  for (let i = 0; i < sessionStorage.length; i += 1) {
                    const key = sessionStorage.key(i);
                    values[key] = sessionStorage.getItem(key);
                  }
                  return values;
                }
                """
            )
        except Exception:
            data = {}
        snapshots.append({"url": url, "sessionStorage": data})
    path.write_text(json.dumps(snapshots, ensure_ascii=False, indent=2), encoding="utf-8")


def _control_html(label: str, state_path: Path) -> str:
    return f"""
    <html lang="ko">
      <head>
        <meta charset="utf-8" />
        <title>{label} 세션 저장</title>
        <style>
          body {{
            font-family: "Malgun Gothic", Arial, sans-serif;
            margin: 32px;
            line-height: 1.55;
            color: #1f2937;
          }}
          button {{
            font-size: 18px;
            padding: 13px 18px;
            border: 0;
            border-radius: 8px;
            background: #174a7c;
            color: white;
            cursor: pointer;
            margin-right: 10px;
            margin-bottom: 10px;
          }}
          button.force {{ background: #6d28d9; }}
          pre {{
            white-space: pre-wrap;
            background: #f3f4f6;
            padding: 14px;
            border-radius: 8px;
            border: 1px solid #e5e7eb;
          }}
          .ok {{ color: #166534; font-weight: 700; }}
          .fail {{ color: #b91c1c; font-weight: 700; }}
          .hint {{ color: #4b5563; }}
        </style>
      </head>
      <body>
        <h1>{label} 로그인 세션 저장</h1>
        <p>첫 번째 탭에서 로그인을 완료한 뒤 이 탭으로 돌아와 아래 버튼을 누르세요.</p>
        <p class="hint">CAPTCHA, 2FA, 회사 선택이 표시되면 사람이 직접 완료해야 합니다. CAPTCHA가 없어도 로그인만 완료되면 저장할 수 있습니다.</p>
        <button onclick="window.requestSave('normal')">로그인 완료, 세션 저장</button>
        <button class="force" onclick="window.requestSave('force')">현재 브라우저 세션 강제 저장</button>
        <h2>진행 상태</h2>
        <pre id="status">준비 중</pre>
        <h2>진단 정보</h2>
        <pre id="diagnostics">저장 경로: {state_path}</pre>
        <h2>결과</h2>
        <pre id="result">아직 저장 전입니다.</pre>
      </body>
    </html>
    """


def _set_status(page: object, message: str) -> None:
    page.evaluate("(message) => { document.querySelector('#status').textContent = message; }", message)


def _append_status(page: object, message: str) -> None:
    page.evaluate(
        "(message) => { const el = document.querySelector('#status'); el.textContent = `${el.textContent}\\n${message}`; }",
        message,
    )


def _show_diagnostics(page: object, signals: LoginSignals, state_path: Path, error: str) -> None:
    page.evaluate(
        """
        ({text}) => {
          document.querySelector('#diagnostics').textContent = text;
        }
        """,
        {"text": _diagnostic_text(signals, state_path, error)},
    )


def _show_success(page: object, message: str, signals: LoginSignals, state_path: Path, session_storage_path: Path) -> None:
    text = (
        f"{message}\n\n"
        f"storage state 저장 파일: {state_path}\n"
        f"sessionStorage 저장 파일: {session_storage_path}\n\n"
        f"{_diagnostic_text(signals, state_path, '')}"
    )
    page.evaluate(
        """
        ({text}) => {
          const el = document.querySelector('#result');
          el.className = 'ok';
          el.textContent = text;
        }
        """,
        {"text": text},
    )


def _show_failure(page: object, signals: LoginSignals, state_path: Path, error: str) -> None:
    text = (
        "세션 저장에 실패했습니다.\n"
        f"{_diagnostic_text(signals, state_path, error)}\n\n"
        "확인 필요: 아직 로그인 페이지에 있거나 회사 선택이 완료되지 않았을 수 있습니다.\n"
        "다음 조치: 첫 번째 탭에서 로그인/회사 선택을 완료한 뒤 '현재 브라우저 세션 강제 저장'을 다시 누르세요."
    )
    page.evaluate(
        """
        ({text}) => {
          const el = document.querySelector('#result');
          el.className = 'fail';
          el.textContent = text;
        }
        """,
        {"text": text},
    )


def _diagnostic_text(signals: LoginSignals, state_path: Path, error: str) -> str:
    matched = ", ".join(signals.matched_signals) if signals.matched_signals else "없음"
    return (
        f"현재 URL: {signals.current_url}\n"
        f"로그인 페이지 여부: {'예' if signals.is_login_page else '아니오'}\n"
        f"쿠키 개수: {signals.cookie_count}\n"
        f"로그인 성공 판정 신호: {matched}\n"
        f"storage state 저장 파일 경로: {state_path}\n"
        f"오류 메시지: {error or '없음'}"
    )


def _signal_summary(signals: LoginSignals) -> str:
    return "|".join(
        [
            signals.current_url,
            str(signals.is_login_page),
            str(signals.cookie_count),
            ",".join(signals.matched_signals),
        ]
    )
