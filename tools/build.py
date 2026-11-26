#!/usr/bin/env python3
"""src/app.html 템플릿(+ src/auth.html 로그인 모듈)에 이미지·사업장·지도 데이터와 라이브러리를 넣어
단독 실행 파일(index.html)과 Claude 아티팩트용 본문(dist/artifact.html)을 만듭니다.
사용법: python3 tools/build.py"""
import json, os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
p = lambda *a: os.path.join(ROOT, *a)
rd = lambda f: open(p(f), encoding="utf-8").read()
t = rd("src/app.html")
img = json.load(open(p("assets/img.json"), encoding="utf-8"))
t = t.replace("/*__LEAFLET_CSS__*/", rd("src/leaflet.css"))
t = t.replace("__CUBE__", img["cube"]).replace("__LOGO__", img["logo"]).replace("__TEAM__", img["team"]).replace("__FIGLOGO__", img["figlogo"]).replace("__POSTER__", img["videoPoster"]).replace("__BRANDSIGN__", img["brandSign"]).replace("__BRANDCASE__", img["brandCase"]).replace("__BRANDOFFICE__", img["brandOffice"]).replace("__SIGFONT__", img["sigFont"]).replace("__APPICON__", img["appIcon"])
import base64
t = t.replace("__VIDEO__", "data:video/mp4;base64," + base64.b64encode(open(p("assets/intro.mp4"),"rb").read()).decode())
t = t.replace("/*__LEAFLET_JS__*/", rd("vendor/leaflet-1.9.4.min.js"))
t = t.replace("/*__EXCELJS__*/", rd("vendor/exceljs-4.4.0.min.js").replace("�", "\\ufffd"))
t = t.replace("/*__QRCODE__*/", rd("vendor/qrcode-generator-1.4.4.js"))
t = t.replace("/*__SHEETJS__*/", rd("vendor/xlsx-0.18.5.core.min.js").replace("\ufffd", "\\ufffd").replace("</script", "<\\/script"))
t = t.replace("/*__IMG__*/{}", json.dumps(img)).replace("/*__SITES__*/{}", rd("assets/sites.json")).replace("/*__ADM__*/{}", rd("assets/adm.json")).replace("/*__GROUP__*/{}", rd("assets/group.json"))
# 제작자 관리(로그인·권한·동기화) 모듈 — src/auth.html + supabase-js
auth = rd("src/auth.html").replace("/*__SUPABASE__*/", rd("vendor/supabase-js-2.119.0.min.js").replace("</script", "<\\/script"))
t = t + "\n" + auth
# 에스텍 통합 로그인 · 프로그램별 사용 권한(app-access) — src/app-access.html
import os
if os.path.exists("src/app-access.html"): t = t + "\n" + rd("src/app-access.html")
# 견적 → 에스텍 제안서 제작시스템 연동(「📑 제안서로 보내기」) — src/proposal-link.html
if os.path.exists(p("src/proposal-link.html")): t = t + "\n" + rd("src/proposal-link.html")
# 갱신 공문(Word·한글) — src/renewal-letter.html + 공문 양식(assets/letter/*.docx·hwpx)
if os.path.exists(p("src/renewal-letter.html")):
    tpl = {k: base64.b64encode(open(p(f"assets/letter/renewal-letter.{k}"),"rb").read()).decode() for k in ("docx","hwpx")}
    t = t + "\n" + rd("src/renewal-letter.html").replace("/*__LETTER_TPL__*/{}", json.dumps(tpl))
# 공통 상단 버튼(stec-bar) · 빈 공간 없는 배치(stec-layout) — 메인 프로그램이 넣어 주는 공통 꼬리 — src/shell-tail.html
if os.path.exists(p("src/shell-tail.html")): t = t + rd("src/shell-tail.html")
os.makedirs(p("dist"), exist_ok=True)
open(p("dist/artifact.html"), "w", encoding="utf-8").write(t)
html = '<!doctype html>\n<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><meta name="theme-color" content="#1b2a86"><link rel="manifest" href="manifest.webmanifest"><link rel="icon" type="image/png" href="favicon.png"><link rel="apple-touch-icon" href="apple-touch-icon.png"><meta name="apple-mobile-web-app-capable" content="yes"><meta name="apple-mobile-web-app-title" content="에스텍 견적"></head><body>\n' + t + '\n</body></html>'
open(p("index.html"), "w", encoding="utf-8").write(html)
print(f"index.html {len(html)//1024} KB")
