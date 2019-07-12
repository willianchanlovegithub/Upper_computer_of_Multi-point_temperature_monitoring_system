Attribute VB_Name = "Module1"
Option Explicit
'窗体最大化定义
Private FormOldWidth  As Long  '保存窗体的原始宽度
Private FormOldHeight  As Long  '保存窗体的原始高度
'kernel32.dll是Windows中非常重要的32位动态链接库文件，属于内核级文件。
'它控制着系统的内存管理、数据的输入输出操作和中断处理
'当Windows启动时，kernel32.dll就驻留在内存中特定的写保护区域，使别的程序无法占用这个内存区域。
Declare Function GetTickCount Lib "kernel32" () As Long

'时间延迟子程序，单位是毫秒(ms)
Public Sub TimeDelay(t As Long)
  Dim TT&
  TT = GetTickCount()
  Do
   DoEvents
  Loop Until GetTickCount() - TT >= t
End Sub

'在调用ReSizeForm前先调用ReSizeInit函数
Public Sub ResizeInit(Form1 As Form)
    Dim Obj  As Control
    FormOldWidth = Form1.ScaleWidth
    FormOldHeight = Form1.ScaleHeight
    On Error Resume Next
    For Each Obj In Form1
    Obj.Tag = Obj.Left & "    " & Obj.Top & "    " & Obj.Width & "    " & Obj.Height & "    "  '双引号内四个空格，多一个少一个都不行
    Next Obj
    On Error GoTo 0
End Sub

'按比例改变表单内各元件的大小,在调用ReSizeForm前先调用ReSizeInit函数
Public Sub ResizeForm(Form1 As Form)
    Dim Pos(4) As Double
    Dim I As Long, TempPos As Long, StartPos As Long
    Dim Obj As Control
    Dim ScaleX As Double, ScaleY As Double
    '在调试时如果出现除数为零的错误，是因为没有设定form初值（既没有在调用ReSizeForm前先调用ReSizeInit函数）
    If FormOldWidth = 0 Then '防止该错误的产生
        Exit Sub
    End If
    ScaleX = Form1.ScaleWidth / FormOldWidth  '保存窗体宽度缩放比例
    ScaleY = Form1.ScaleHeight / FormOldHeight  '保存窗体高度缩放比例
    On Error Resume Next
    For Each Obj In Form1
        StartPos = 1
        For I = 0 To 4
            '读取控件的原始位置与大小
            TempPos = InStr(StartPos, Obj.Tag, "    ", vbTextCompare)  '双引号内四个空格，多一个少一个都不行
            If TempPos > 0 Then
                Pos(I) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                StartPos = TempPos + 1
            Else
                Pos(I) = 0
            End If
            '根据控件的原始位置及窗体改变大小的比例对控件重新定位与改变大小
            Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
        Next I
    Next Obj
    On Error GoTo 0
End Sub

