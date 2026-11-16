#!/usr/bin/env python3
"""Builds index.html (the standalone, installable app) from app.html.

app.html is the single source: it is published as-is as the Claude artifact, where the
platform supplies the document shell and storage. This script wraps the same content in a
full HTML document that loads Firebase, so the app also runs on its own at a plain URL.
Run after every change to app.html:  python3 build.py
"""
import json, pathlib

here = pathlib.Path(__file__).parent
app = (here / "app.html").read_text(encoding="utf-8")

# Firebase web config. These values identify the project and are public by design;
# access is controlled by the Firestore security rules, not by keeping them secret.
FIREBASE_CONFIG = {
    "apiKey": "AIzaSyAfgFLi09nUFYIx-ZjrIx6RpFu_JBW7Lyk",
    "authDomain": "swipe-anki.firebaseapp.com",
    "projectId": "swipe-anki",
    "storageBucket": "swipe-anki.firebasestorage.app",
    "messagingSenderId": "394041819232",
    "appId": "1:394041819232:web:3c2fd2a60ee92e47650305",
}

head = """<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>스와이프 암기장</title>
<meta name="theme-color" content="#f1f0f6" media="(prefers-color-scheme: light)">
<meta name="theme-color" content="#131119" media="(prefers-color-scheme: dark)">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="암기장">
<link rel="manifest" href="manifest.webmanifest">
<link rel="icon" href="icon-192.png">
<link rel="apple-touch-icon" href="icon-180.png">
<style>:root{color-scheme:light;padding:env(safe-area-inset-top,0px) 0 env(safe-area-inset-bottom,0px)}</style>
<script src="vendor/firebase-app-compat.js"></script>
<script src="vendor/firebase-auth-compat.js"></script>
<script src="vendor/firebase-firestore-compat.js"></script>
<script>
window.SC_FIREBASE_CONFIG = %s;
if ('serviceWorker' in navigator) addEventListener('load', function () { navigator.serviceWorker.register('sw.js').catch(function () {}); });
</script>
</head>
<body>
""" % json.dumps(FIREBASE_CONFIG, indent=2)

(here / "index.html").write_text(head + app + "</body>\n</html>\n", encoding="utf-8")
print("wrote index.html")
