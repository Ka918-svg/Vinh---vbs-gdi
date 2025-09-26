' gdi_ellipse.vbs - Spam ellipse màu ngẫu nhiên (dễ thương, vô hại)
Option Explicit

Randomize
Dim x, y, w, h, r, g, b
Dim obj, desktop

' Tạo đối tượng vẽ GDI trên desktop
Set obj = CreateObject("WScript.Shell")
Set desktop = GetObject("winmgmts:root\cimv2:Win32_Desktop")

MsgBox "Continue?", vbInformation, "vinh"

Do
    ' Vị trí và kích thước ellipse
    x = Int(Rnd * 1000)
    y = Int(Rnd * 700)
    w = Int(Rnd * 300) + 50
    h = Int(Rnd * 300) + 50

    ' Màu ngẫu nhiên
    r = Int(Rnd * 255)
    g = Int(Rnd * 255)
    b = Int(Rnd * 255)

    ' Vẽ ellipse bằng PowerShell + GDI+
    obj.Run "powershell -command " & _
        """Add-Type -AssemblyName System.Drawing; " & _
        "$bmp = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero); " & _
        "$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(" & r & "," & g & "," & b & ")); " & _
        "$bmp.DrawEllipse($pen," & x & "," & y & "," & w & "," & h & "); " & _
        "$pen.Dispose(); $bmp.Dispose()""", 0, True

    WScript.Sleep 100
Loop
