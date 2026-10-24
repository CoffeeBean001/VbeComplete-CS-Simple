@echo off
rem One-click self check. Nine steps, cheapest first:
rem   1 pure logic           no COM, no desktop
rem   2 snippet store        temp db, your own snippets are untouched
rem   3 offscreen UI preview two PNGs you can eyeball
rem   4 icon + autostart     embedded app.ico + HKCU Run entry, restored afterwards
rem   5 candidate rules      runs the engine against the REAL module text
rem   6 parens on accept     needs the VBE code window open
rem   7 caret landing        needs the VBE code window open
rem   8 hotkey commands      needs the VBE code window open
rem   9 enter auto-close     needs the VBE code window open and drives the REAL hook
rem
rem Steps 5-8 say "no code module, skip" instead of failing when the code
rem window is closed -- there is no way to reach a code pane from COM then.
rem
rem NOTE: CRLF line endings, and no bare parenthesis in echo text.
setlocal
set "EXE=%~dp0dist\VbeComplete.exe"

if not exist "%EXE%" goto nobuild

echo.
echo [check] 1/9 pure logic - candidate pool, rules, scope, snippets
"%EXE%" --selftest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0selftest.log' -Tail 4"
echo.

echo [check] 2/9 snippet store - temp db, the real one is untouched
"%EXE%" --snipetest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0snipetest.log' -Tail 4"
echo.

echo [check] 3/9 offscreen UI preview - popup-preview.png / snippet-preview.png
"%EXE%" --uitest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0uitest.log' -Tail 4"
echo.

echo [check] 4/9 icon + autostart - embedded app.ico, Run entry restored afterwards
"%EXE%" --startuptest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0startuptest.log' -Tail 8"
echo.

echo [check] 5/9 candidate rules - needs the VBE code window open
"%EXE%" --enginetest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0enginetest.log' -Tail 4"
echo.

echo [check] 6/9 parens after accept - needs the VBE code window open
"%EXE%" --paretest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0paretest.log' -Tail 6"
echo.

echo [check] 7/9 caret landing - needs the VBE code window open
"%EXE%" --carettest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0carettest.log' -Tail 4"
echo.

echo [check] 8/9 hotkey commands - needs the VBE code window open
"%EXE%" --hktest
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0hktest.log' -Tail 4"
echo.
echo [check] 9/9 enter auto-close - drives the REAL keyboard hook
"%EXE%" --enterprobe
powershell -NoProfile -Command "Get-Content -LiteralPath '%~dp0enterprobe.log' -Tail 8"
echo.
echo Logs: selftest.log / snipetest.log / uitest.log / startuptest.log /
echo       enginetest.log / paretest.log / carettest.log / hktest.log /
echo       enterprobe.log
pause
endlocal
exit /b 0

:nobuild
echo [check] dist\VbeComplete.exe not found - run build.bat first.
pause
endlocal
exit /b 1
