#!/usr/bin/env python3
"""Browser-only regression checks. Mock bridge; does not test C#/WPF/WebView2.
Requires: Python 3, playwright, Chromium. Uses synthetic records only.
"""
import argparse
import datetime
import json
import pathlib
import platform
import re
import sys
from playwright.sync_api import sync_playwright

SCREEN = {
    "card": {"width": 840, "gap": 12, "padding": [16, 16, 16, 16],
             "font": "Arial", "fontSize": 11, "keyValueFontSize": 22,
             "judgmentFontSize": 18, "unsearchedFontSize": 16},
    "sections": [{"type": "titleBar", "brand": "RDV regression", "tags": [], "buttons": []},
                 {"type": "keyPanel", "title": "Search", "label": "Key", "value": "keyLabel",
                  "inputLabel": "Search", "inputWidth": 180, "maxLength": 8,
                  "buttons": [{"action": "search", "text": "Search"},
                              {"action": "clear", "text": "Clear"},
                              {"action": "workState", "text": "Mark"},
                              {"action": "updateRecords", "text": "Job one", "job": "one"},
                              {"action": "updateRecords", "text": "Job two", "job": "two"}]}]}
STATE = {"opsEnabled": True, "workEnabled": True, "key": "", "values": {},
         "workText": "Mark", "pending": 0, "judgments": {}}
BRIDGE = """window.rdvTestMessages=[]; window.chrome=window.chrome||{};
window.chrome.webview={postMessage:m=>window.rdvTestMessages.push(m),
addEventListener:(name,h)=>{window.rdvDeliver=m=>h({data:m});}};"""


def check(condition, detail):
    if not condition:
        raise AssertionError(detail)


def clear(page):
    page.evaluate("window.rdvTestMessages=[]")


def messages(page, kind):
    return page.evaluate("t=>window.rdvTestMessages.filter(m=>m.type===t)", kind)


def modal(page, kind, content):
    page.evaluate("m=>window.rdvDeliver(m)",
                  {"type": "modalOpen", "token": 123, "modal": kind, "content": content})
    page.wait_for_selector(".veil.show")
    page.wait_for_timeout(30)
    clear(page)


def press_enter(page, composing=False, selector="#input"):
    page.locator(selector).evaluate("(el, composing)=>el.dispatchEvent(new KeyboardEvent('keydown',"
                                   "{key:'Enter',code:'Enter',keyCode:composing?229:13,"
                                   "isComposing:composing,bubbles:true,cancelable:true}))", composing)


def input_normalization(page):
    page.locator("#input").evaluate("el=>{el.innerHTML='<div>001\\n002</div>';el.dispatchEvent(new InputEvent('input',{bubbles:true}));}")
    check(page.locator("#input").text_content() == "001002", "visible newlines were not removed")
    check(page.locator("#input").evaluate("el=>el.children.length") == 0, "rich DOM remained")
    check(messages(page, "key")[-1]["value"] == "001002", "visible/input values differ")


def ime_main(page):
    press_enter(page, True)
    check(not messages(page, "action"), "IME confirmation incorrectly triggered search")


def enter_once(page):
    page.locator("#input").evaluate("el=>el.textContent='001'")
    press_enter(page)
    actions = messages(page, "action")
    check(len(actions) == 1 and actions[0]["name"] == "search" and actions[0]["key"] == "001", str(actions))


def max_length(page):
    page.locator("#input").evaluate("el=>{el.textContent='123456789012';el.dispatchEvent(new InputEvent('input',{bubbles:true}));}")
    check(page.locator("#input").text_content() == "12345678", "input exceeded configured length")


def duplicate_actions(page):
    buttons = page.locator('[data-action="updateRecords"]')
    ids = buttons.evaluate_all("els=>els.map(el=>el.id)")
    check(len(ids) == 2 and len(set(ids)) == 2, "duplicate DOM ids: " + str(ids))
    buttons.nth(0).click(); buttons.nth(1).click()
    check([m["job"] for m in messages(page, "action")] == ["one", "two"], "job routing changed")


def disabled_actions(page):
    page.evaluate("s=>window.rdvBridge.state(s)", dict(STATE, opsEnabled=False))
    page.locator("#b-upd").evaluate("el=>el.click()")
    press_enter(page)
    check(not messages(page, "action"), "disabled action fired")
    check(page.locator("#input").get_attribute("contenteditable") == "false", "disabled input remained editable")


def ime_modal(page):
    modal(page, "confirm", {"title": "Confirmation", "body": "test", "ask": True})
    press_enter(page, True, ".veil.show [data-modal-default=true]")
    check(not messages(page, "modalResult"), "IME confirmation committed dialog")
    press_enter(page, False, ".veil.show [data-modal-default=true]")
    result = messages(page, "modalResult")
    check(len(result) == 1 and result[0]["result"]["ok"], str(result))


def unmatched(page, discard):
    modal(page, "unmatched", {"title": "Conflicts", "body": "Keep or discard local only", "discardText": "Discard local",
                              "rows": [["1", "<img src=x onerror=alert(1)>", "state-conflict"]]})
    check(page.locator(".veil.show img").count() == 0, "record text became HTML")
    check("<img" in page.locator(".veil.show").text_content(), "literal text lost")
    if discard:
        page.locator(".veil.show .foot .btn").filter(has_text="Discard local").click()
    else:
        press_enter(page, False, ".veil.show [data-modal-default=true]")
    result = messages(page, "modalResult")
    check(len(result) == 1 and result[0]["result"].get("discard") is discard, str(result))


def escape_keeps(page):
    modal(page, "unmatched", {"title": "Conflicts", "body": "Test", "rows": [["1", "001", "changed"]]})
    page.keyboard.press("Escape")
    result = messages(page, "modalResult")
    check(len(result) == 1 and not result[0]["result"]["ok"] and not result[0]["result"].get("discard"), str(result))


def export(page, safe=True, extension="csv"):
    modal(page, "export", {"title": "Export", "fields": [{"ref": "T.id", "label": "ID", "kind": "text"}],
                           "defaults": ["T.id"], "destination": "C:\\temp\\output." + extension})
    checkbox = page.locator("#export-excel-safe")
    check(checkbox.is_checked(), "formula-safe default not checked")
    checkbox.set_checked(safe)
    page.locator(".veil.show [data-modal-default=true]").click()
    result = messages(page, "modalResult")
    if extension != "csv":
        check(not result and page.locator(".veil.show .setting-error").is_visible(), "non-CSV destination accepted")
    else:
        check(len(result) == 1 and result[0]["result"]["excelSafe"] is safe and
              result[0]["result"]["fields"] == ["T.id"], str(result))


def rerender_ids(page):
    page.evaluate("s=>window.rdvBridge.render(s)", SCREEN)
    check(page.locator("#b-upd").count() == 1 and page.locator("#b-upd-2").count() == 1, "id counters not reset")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=pathlib.Path, default=pathlib.Path(__file__).resolve().parents[1])
    parser.add_argument("--chromium", default=None, help="Optional installed Chromium executable path")
    parser.add_argument("--output", type=pathlib.Path)
    parser.add_argument("--baseline", action="store_true", help="Run only three original-regression reproductions; failures are expected")
    args = parser.parse_args()
    tests = [("input-visible-value-normalization", input_normalization), ("ime-enter-does-not-search", ime_main),
             ("duplicate-actions-unique-id-and-routing", duplicate_actions)]
    if not args.baseline:
        tests += [("enter-search-exactly-once", enter_once), ("input-max-length", max_length),
                  ("disabled-actions-blocked", disabled_actions), ("ime-modal-does-not-commit", ime_modal),
                  ("unmatched-default-keeps-and-escapes-html", lambda p: unmatched(p, False)),
                  ("unmatched-explicit-discard", lambda p: unmatched(p, True)), ("escape-keeps-pending", escape_keeps),
                  ("export-excel-safe-default", lambda p: export(p, True)),
                  ("export-raw-option", lambda p: export(p, False)),
                  ("export-rejects-non-csv", lambda p: export(p, True, "xlsx")),
                  ("rerender-resets-action-ids", rerender_ids)]
    output = {"scope": "Browser JS/DOM with synthetic WebView bridge; NOT C#/WPF/SMB",
              "baseline_reproduction": args.baseline, "utc": datetime.datetime.now(datetime.timezone.utc).isoformat(),
              "platform": platform.platform(), "tests": []}
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        output["browser"] = browser.version
        for name, action in tests:
            page = browser.new_page(viewport={"width": 1100, "height": 900})
            errors = []
            page.on("pageerror", lambda error: errors.append(str(error)))
            page.add_init_script(BRIDGE)
            try:
                # Inject the exact bundled files without a network/server or file:// permissions.
                web = args.root.resolve() / "web"
                html = (web / "index.html").read_text(encoding="utf-8-sig")
                html = re.sub(r'<script\b[^>]*\bsrc=[^>]*>\s*</script>', '', html, flags=re.I)
                html = re.sub(r'<link\b[^>]*\brel=["\']stylesheet["\'][^>]*>', '', html, flags=re.I)
                page.set_content(html)
                page.evaluate(BRIDGE)
                page.add_style_tag(content=(web / "app.css").read_text(encoding="utf-8-sig"))
                page.add_script_tag(content=(web / "app.js").read_text(encoding="utf-8-sig"))
                page.wait_for_function("!!window.rdvBridge")
                page.evaluate("s=>window.rdvBridge.render(s)", SCREEN)
                page.evaluate("s=>window.rdvBridge.state(s)", STATE)
                page.wait_for_timeout(30)
                clear(page)
                action(page)
                page.wait_for_timeout(20)
                check(not errors, "Uncaught browser error: " + str(errors))
                output["tests"].append({"name": name, "status": "PASS"})
            except Exception as exc:
                output["tests"].append({"name": name, "status": "FAIL", "detail": str(exc), "page_errors": errors})
            finally:
                page.close()
            print(output["tests"][-1]["status"], name, output["tests"][-1].get("detail", ""), flush=True)
        browser.close()
    output["passed"] = sum(t["status"] == "PASS" for t in output["tests"])
    output["failed"] = len(tests) - output["passed"]
    destination = args.output or args.root / "tests" / "results" / "web-results.json"
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return 1 if output["failed"] else 0


if __name__ == "__main__":
    sys.exit(main())
