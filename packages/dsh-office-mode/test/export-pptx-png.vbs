' 用本机 PowerPoint 把 pptx 每页导出成 PNG，供人工查看版面。
' 跑法：cscript //nologo export-pptx-png.vbs <pptx> <outDir> [width] [height]
Option Explicit

Dim args, src, outDir, width, height
Dim fso, app, pres, i, n
Set args = WScript.Arguments
If args.Count < 2 Then
    WScript.Echo "usage: export-pptx-png.vbs <pptx> <outDir> [width] [height]"
    WScript.Quit 2
End If

src = args(0)
outDir = args(1)
width = 1600
height = 900
If args.Count >= 4 Then
    width = CLng(args(2))
    height = CLng(args(3))
End If

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(outDir) Then fso.CreateFolder(outDir)

Set app = CreateObject("PowerPoint.Application")
' ReadOnly:=True, Untitled:=False, WithWindow:=False
Set pres = app.Presentations.Open(src, True, False, False)

n = 0
For i = 1 To pres.Slides.Count
    Dim name
    If i < 10 Then
        name = "slide-0" & i & ".png"
    Else
        name = "slide-" & i & ".png"
    End If
    pres.Slides(i).Export outDir & "\" & name, "PNG", width, height
    n = n + 1
Next

pres.Close
app.Quit
WScript.Echo "exported " & n & " slides to " & outDir
