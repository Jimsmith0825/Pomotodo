@echo off
rem Build PomoTodo.exe with the C# compiler built into Windows (.NET Framework 4.x)
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
"%CSC%" /nologo /target:winexe /optimize+ /win32icon:tomato.ico /out:PomoTodo.exe /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.IO.Compression.dll /r:System.Xml.Linq.dll /r:System.Xml.dll /r:System.Core.dll PomoTodo.cs
pause
