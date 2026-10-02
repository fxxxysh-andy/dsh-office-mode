' Open the generated .docx/.xlsx/.pptx with the real Microsoft Office installed here,
' to prove the files are valid rather than merely parseable by our own reader.
'
' Why VBScript rather than PowerShell:
'   1. pwsh 7's New-Object -ComObject cannot bind members on this machine (the type
'      library fails to load); plain IDispatch through cscript works.
'   2. PowerPoint fails on any file with media when opened WithWindow:=False
'      (Office's own templates and python-pptx output fail the same way), so
'      WithWindow must be True.
'   3. cscript's stdout is ANSI, which mangles CJK file names, so the result is
'      written to a UTF-8 report file instead.
'
' This file is deliberately pure ASCII: cscript reads .vbs as ANSI unless the file
' carries a UTF-8 BOM, so non-ASCII source would be misparsed.
'
' Usage: cscript //nologo test\validate-com.vbs <directory> <report-file>

Option Explicit

Dim args, target, report, fso, folder, file, ok, bad, appCache, ext, lines

Set args = WScript.Arguments
If args.Count < 2 Then
    WScript.Echo "usage: cscript //nologo validate-com.vbs <directory> <report-file>"
    WScript.Quit 2
End If
target = args(0)
report = args(1)

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(target) Then
    WScript.Echo "no such directory: " & target
    WScript.Quit 2
End If

Set appCache = CreateObject("Scripting.Dictionary")
Set folder = fso.GetFolder(target)
ok = 0
bad = 0
lines = ""

For Each file In folder.Files
    ext = LCase(fso.GetExtensionName(file.Name))
    If ext = "docx" Or ext = "xlsx" Or ext = "pptx" Then
        Dim failed, detail
        failed = False
        detail = ""
        On Error Resume Next
        Err.Clear

        If ext = "docx" Then
            Dim word, doc
            Set word = AppFor("word", "Word.Application")
            Set doc = word.Documents.Open(file.Path, False, True, False)
            If Err.Number <> 0 Then
                failed = True
                detail = Err.Description
            Else
                detail = "pages=" & doc.ComputeStatistics(2) & " tables=" & doc.Tables.Count & " paras=" & doc.Paragraphs.Count
                If Err.Number <> 0 Then detail = "(statistics failed: " & Err.Description & ")"
                doc.Close 0
            End If
        ElseIf ext = "xlsx" Then
            Dim excel, wb, charts
            Set excel = AppFor("excel", "Excel.Application")
            ' Call with a single argument on purpose. Through pure IDispatch, passing
            ' the optional UpdateLinks/ReadOnly positionally makes Excel answer
            ' "cannot get the Open property of class Workbooks" -- a binding
            ' complaint that looks exactly like a broken file but is not one.
            ' Verified with a control probe: Open(<missing path>) then reports a
            ' normal 1004 file error, and Workbooks.Add succeeds.
            Set wb = excel.Workbooks.Open(file.Path)
            If Err.Number <> 0 Then
                failed = True
                detail = Err.Description
            Else
                charts = 0
                Dim si
                For si = 1 To wb.Worksheets.Count
                    charts = charts + wb.Worksheets.Item(si).ChartObjects.Count
                    If Err.Number <> 0 Then
                        charts = -1
                        Err.Clear
                        Exit For
                    End If
                Next
                detail = "sheets=" & wb.Worksheets.Count & " first=" & wb.Worksheets.Item(1).Name
                If charts >= 0 Then detail = detail & " charts=" & charts
                If Err.Number <> 0 Then detail = "(statistics failed: " & Err.Description & ")"
                wb.Close False
            End If
        ElseIf ext = "pptx" Then
            Dim ppt, pres, slides, shape, chartShapes
            Set ppt = AppFor("ppt", "PowerPoint.Application")
            ' -1 = msoTrue: WithWindow has to be true or media-bearing files refuse to open.
            Set pres = ppt.Presentations.Open(file.Path, -1, 0, -1)
            If Err.Number <> 0 Then
                failed = True
                detail = Err.Description
            Else
                slides = pres.Slides.Count
                chartShapes = 0
                If Err.Number <> 0 Then
                    detail = "(Slides.Count failed: " & Err.Description & ")"
                Else
                    Dim pi, qi
                    For pi = 1 To slides
                        For qi = 1 To pres.Slides.Item(pi).Shapes.Count
                            Set shape = pres.Slides.Item(pi).Shapes.Item(qi)
                            If shape.HasChart Then chartShapes = chartShapes + 1
                            If Err.Number <> 0 Then
                                chartShapes = -1
                                Err.Clear
                                Exit For
                            End If
                        Next
                        If chartShapes < 0 Then Exit For
                    Next
                    detail = "slides=" & slides
                    If chartShapes >= 0 Then detail = detail & " chartShapes=" & chartShapes
                End If
                pres.Close
            End If
        End If
        On Error GoTo 0

        If failed Then
            bad = bad + 1
            lines = lines & "FAIL  " & file.Name & "  -> " & detail & vbCrLf
        Else
            ok = ok + 1
            lines = lines & "OK    " & file.Name & "  -> " & detail & vbCrLf
        End If
    End If
Next

Dim key
For Each key In appCache.Keys
    On Error Resume Next
    appCache(key).Quit
    On Error GoTo 0
Next

lines = lines & "com-validate: " & ok & " passed / " & bad & " failed" & vbCrLf

Dim stream
Set stream = CreateObject("ADODB.Stream")
stream.Type = 2
stream.Charset = "utf-8"
stream.Open
stream.WriteText lines
stream.SaveToFile report, 2
stream.Close

WScript.Echo "report written to " & report
If bad > 0 Then
    WScript.Quit 1
Else
    WScript.Quit 0
End If

' Reuse one application instance per host: starting Word/Excel/PowerPoint per file
' is slow, and a fresh PowerPoint per file can leave orphan processes behind.
Function AppFor(cacheKey, progId)
    If appCache.Exists(cacheKey) Then
        Set AppFor = appCache(cacheKey)
    Else
        Dim created
        Set created = CreateObject(progId)
        appCache.Add cacheKey, created
        Set AppFor = created
    End If
End Function
