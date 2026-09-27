' Step-by-step Word COM probe.
'
' Why this exists: on 2026-09-22 a Word COM failure was misdiagnosed as "Word COM
' validation is unavailable on this machine" and written into the docs as a known
' limitation. The real cause was orphaned WINWORD/EXCEL processes left over from the
' previous day: they held COM hostage, so calls hung and reported "Word failed to
' raise an event". Killing them made the very same file pass immediately.
'
' validate-com.vbs reports one line per file, which cannot distinguish
'   (a) Word refuses to start,
'   (b) Word starts but refuses this file,
'   (c) Word and the file are both fine but COM is wedged by a stale process.
' This probe reports the error code at each step, so those three cases are separable.
' All-zero across the six steps means Word and the file are both fine.
'
' Usage: cscript //nologo test\probe-word-com.vbs <file.docx> <report.txt>
'
' Pure ASCII on purpose: cscript reads .vbs as ANSI unless the file carries a UTF-8
' BOM, so non-ASCII source would be misparsed.

Option Explicit

Dim args, target, report, fso, word, doc, lines, step

Set args = WScript.Arguments
If args.Count < 2 Then
    WScript.Echo "usage: cscript //nologo probe-word-com.vbs <file.docx> <report.txt>"
    WScript.Quit 2
End If
target = args(0)
report = args(1)

Set fso = CreateObject("Scripting.FileSystemObject")
lines = ""

If Not fso.FileExists(target) Then
    WScript.Echo "no such file: " & target
    WScript.Quit 2
End If

' An absolute path matters here. cscript's working directory is system32, so a
' relative path is resolved against it and Word reports "the file could not be
' found" -- which looks like a broken document but is only a path problem.
If InStr(target, ":") = 0 Then
    lines = lines & "WARNING: target is not an absolute path; Word resolves relative" & vbCrLf
    lines = lines & "         paths against cscript's cwd (system32). Pass a full path." & vbCrLf
End If

On Error Resume Next
Err.Clear

step = "CreateObject(Word.Application)"
Set word = CreateObject("Word.Application")
lines = lines & "[" & step & "] err=" & Err.Number & " " & Err.Description & vbCrLf

If Err.Number = 0 Then
    step = "word.Visible = False"
    Err.Clear
    word.Visible = False
    lines = lines & "[" & step & "] err=" & Err.Number & " " & Err.Description & vbCrLf

    step = "word.DisplayAlerts = 0"
    Err.Clear
    word.DisplayAlerts = 0
    lines = lines & "[" & step & "] err=" & Err.Number & " " & Err.Description & vbCrLf

    step = "word.Version"
    Err.Clear
    Dim ver
    ver = word.Version
    lines = lines & "[" & step & "] err=" & Err.Number & " " & Err.Description & " value=" & ver & vbCrLf

    ' Positional args only: named arguments are unreliable through pure IDispatch
    ' (same reason validate-com.vbs passes them positionally).
    step = "Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)"
    Err.Clear
    Set doc = word.Documents.Open(target, False, True, False)
    lines = lines & "[" & step & "] err=" & Err.Number & " " & Err.Description & vbCrLf

    If Err.Number = 0 Then
        step = "doc.Paragraphs.Count"
        Err.Clear
        Dim n
        n = doc.Paragraphs.Count
        lines = lines & "[" & step & "] err=" & Err.Number & " " & Err.Description & " value=" & n & vbCrLf

        step = "doc.Close(0)"
        Err.Clear
        doc.Close 0
        lines = lines & "[" & step & "] err=" & Err.Number & " " & Err.Description & vbCrLf
    End If

    step = "word.Quit"
    Err.Clear
    word.Quit
    lines = lines & "[" & step & "] err=" & Err.Number & " " & Err.Description & vbCrLf
End If

On Error GoTo 0

' If anything failed, point at the most common cause before anything else.
lines = lines & vbCrLf
lines = lines & "If a step failed, first kill leftover Office processes:" & vbCrLf
lines = lines & "  Get-Process WINWORD,EXCEL,POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force" & vbCrLf
lines = lines & "Stale Office processes hold COM hostage: calls hang and report" & vbCrLf
lines = lines & """Word failed to raise an event"", which is easily misread as" & vbCrLf
lines = lines & """this machine cannot do COM validation""." & vbCrLf

Dim stream
Set stream = CreateObject("ADODB.Stream")
stream.Type = 2
stream.Charset = "utf-8"
stream.Open
stream.WriteText lines
stream.SaveToFile report, 2
stream.Close

WScript.Echo "probe report: " & report
