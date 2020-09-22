<div align="center">

## Changing behaviour of ComboBox and ListBox \(Class\)


</div>

### Description

Changing behaviour of ComboBox and ListBox (Class)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chong Long Choo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chong-long-choo.md)
**Level**          |Unknown
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chong-long-choo-changing-behaviour-of-combobox-and-listbox-class__1-3507/archive/master.zip)

### API Declarations

```
Option Explicit
' Name:   Changing behaviour of ComboBox and ListBox
' Author:  Chong Long Choo
' Email: chonglongchoo@hotmail.com
' Date:   14 September 1999
'<--------------------------Disclaimer------------------------------->
'
'This sample is free. You can use the sample in any form. Use this
'sample at your own risk! I have no warranty for this sample.
'
'<--------------------------Disclaimer------------------------------->
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageLongByRef Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function LBItemFromPt Lib "COMCTL32.DLL" (ByVal hLB As Long, ByVal ptX As Long, ByVal ptY As Long, ByVal bAutoScroll As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETITEMHEIGHT = &H154
Private Const CB_GETLBTEXTLEN = &H149
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const ANSI_FIXED_FONT = 11
Private Const ANSI_VAR_FONT = 12
Private Const SYSTEM_FONT = 13
Private Const DEFAULT_GUI_FONT = 17 'win95/98 only
Private Const SM_CXHSCROLL = 21
Private Const SM_CXHTHUMB = 10
Private Const SM_CXVSCROLL = 2
Private Const DT_CALCRECT = &H400
Private Const LB_SETCURSEL = &H186
Private Const LB_GETCURSEL = &H188
Private Const LB_SETTABSTOPS = &H192
Private Const WM_USER = &H400
Private Const LB_SETHORIZONTALEXTENT = WM_USER + 21
Private Const tmp = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Private Type SIZE
  cX As Long
  cY As Long
End Type
Private Type POINTAPI
  X As Long
  Y As Long
End Type
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
```


### Source Code

```
' Adjust Drop Down Width (ComboBox)
Public Sub AdjDropDownWidth(ByVal NewDropDownWidth As Long, ByVal ComboHwnd As Long)
  Call SendMessageLong(ComboHwnd, CB_SETDROPPEDWIDTH, NewDropDownWidth, 0)
  Call SendMessageLong(ComboHwnd, CB_GETDROPPEDWIDTH, 0, 0)
End Sub
Private Function GetCmbItemWidth(ByVal FormHwnd As Long) As Long
  Dim hFont As Long
  Dim hFontOld As Long
  Dim r As Long
  Dim avgWidth As Long
  Dim hDC As Long
  Dim sz As SIZE
  hDC = GetDC(FormHwnd)
  hFont = GetStockObject(ANSI_VAR_FONT)
  hFontOld = SelectObject(hDC, hFont)
  Call GetTextExtentPoint32(hDC, tmp, 52, sz)
  avgWidth = (sz.cX / 52)
  Call SelectObject(hDC, hFontOld)
  Call DeleteObject(hFont)
  Call ReleaseDC(FormHwnd, hDC)
  GetCmbItemWidth = avgWidth
End Function
' Set Drop Down Height (ComboBox)
Public Sub SetCmbDropDownHeight(ByVal numItemsToDisplay As Byte, ByVal objCombo As ComboBox)
  Dim cWidth As Long
  Dim newHeight As Long
  Dim oldScaleMode As Long
  Dim itemHeight As Long
  Dim ComboHwnd As Long
  ComboHwnd = objCombo.hwnd
  oldScaleMode = objCombo.Parent.ScaleMode
  objCombo.Parent.ScaleMode = vbPixels
  cWidth = objCombo.Width
  itemHeight = SendMessageLong(ComboHwnd, CB_GETITEMHEIGHT, 0, 0)
  newHeight = itemHeight * (numItemsToDisplay + 2)
  Call MoveWindow(ComboHwnd, objCombo.Left / Screen.TwipsPerPixelX, objCombo.Top / Screen.TwipsPerPixelX, objCombo.Width / Screen.TwipsPerPixelX, newHeight, True)
  objCombo.Parent.ScaleMode = oldScaleMode
End Sub
' Auto Adjust Drop Down Width (ComboBox)
Public Sub AutoAdjCombo(ByVal objCombo As ComboBox)
  Dim i As Long
  Dim NumOfChars As Long
  Dim LongestComboItem As Long
  Dim avgCharWidth As Long
  Dim NewDropDownWidth As Long
  Dim ComboHwnd As Long
  ComboHwnd = objCombo.hwnd
  For i = 0 To objCombo.ListCount - 1
    NumOfChars = SendMessageLong(ComboHwnd, CB_GETLBTEXTLEN, i, 0)
    If NumOfChars > LongestComboItem Then LongestComboItem = NumOfChars
  Next
  avgCharWidth = GetCmbItemWidth(objCombo.Parent.hwnd)
  NewDropDownWidth = (LongestComboItem - 2) * avgCharWidth
  Call SendMessageLong(ComboHwnd, CB_SETDROPPEDWIDTH, NewDropDownWidth, 0)
  Call SendMessageLong(ComboHwnd, CB_GETDROPPEDWIDTH, 0, 0)
End Sub
' Show Drop Down (ComboBox)
Public Sub Dropdown(ByVal ComboHwnd As Long)
  Call SendMessageLong(ComboHwnd, CB_SHOWDROPDOWN, True, 0)
End Sub
' Hide Drop Down (ComboBox)
Public Sub HideDropDown(ComboHwnd As Long)
  Call SendMessageLong(ComboHwnd, CB_SHOWDROPDOWN, False, ByVal 0)
End Sub
' Copy contents of a listbox to another listbox
Public Function CopyListToList(SourceHwnd As Long, DestHwnd As Long) As Long
  Dim c As Long
  Const LB_GETCOUNT = &H18B
  Const LB_GETTEXT = &H189
  Const LB_ADDSTRING = &H180
  Dim numitems As Long
  Dim sItemText As String * 255
  numitems = SendMessageLong(SourceHwnd, LB_GETCOUNT, 0&, 0&)
  LockWinUpdate DestHwnd
  If numitems > 0 Then
    For c = 0 To numitems - 1
      Call SendMessageStr(SourceHwnd, LB_GETTEXT, c, ByVal sItemText)
      Call SendMessageStr(DestHwnd, LB_ADDSTRING, 0&, ByVal sItemText)
    Next
  End If
  LockWinUpdate 0&
  numitems = SendMessageLong(DestHwnd, LB_GETCOUNT, 0&, 0&)
  CopyListToList = numitems
End Function
' Copy contents of a listbox to a combobox
Public Function CopyListToCombo(SourceHwnd As Long, DestHwnd As Long) As Long
  Dim c As Long
  Const LB_GETCOUNT = &H18B
  Const LB_GETTEXT = &H189
  Const CB_GETCOUNT = &H146
  Const CB_ADDSTRING = &H143
  Dim numitems As Long
  Dim sItemText As String * 255
  numitems = SendMessageLong(SourceHwnd, LB_GETCOUNT, 0&, 0&)
  LockWinUpdate DestHwnd
  If numitems > 0 Then
    For c = 0 To numitems - 1
      Call SendMessageStr(SourceHwnd, LB_GETTEXT, c, ByVal sItemText)
      Call SendMessageStr(DestHwnd, CB_ADDSTRING, 0&, ByVal sItemText)
    Next
  End If
  LockWinUpdate 0&
  numitems = SendMessageLong(DestHwnd, CB_GETCOUNT, 0&, 0&)
  CopyListToCombo = numitems
End Function
'Set horizontal extent (ListBox)
Public Sub SetLBHorizontalExtent(objLB As ListBox)
  Dim i As Integer
  Dim res As Long
  Dim Scrollwidth As Long
  With objLB
    For i = 0 To .ListCount
      If .Parent.TextWidth(.List(i)) > Scrollwidth Then _
      Scrollwidth = .Parent.TextWidth(.List(i))
    Next i
    res = SendMessage(.hwnd, LB_SETHORIZONTALEXTENT, _
      (Scrollwidth + 100) / Screen.TwipsPerPixelX, 0)
  End With
End Sub
' Highlight An Item When Your Mouse Is Over It (ListBox)
Public Sub HighlightLBItem(ByVal LBHwnd As Long, ByVal X As Single, ByVal Y As Single)
  Dim ItemIndex As Long
  Dim AtThisPoint As POINTAPI
  AtThisPoint.X = X \ Screen.TwipsPerPixelX
  AtThisPoint.Y = Y \ Screen.TwipsPerPixelY
  Call ClientToScreen(LBHwnd, AtThisPoint)
  ItemIndex = LBItemFromPt(LBHwnd, AtThisPoint.X, AtThisPoint.Y, False)
  If ItemIndex <> SendMessage(LBHwnd, LB_GETCURSEL, 0, 0) Then
    Call SendMessage(LBHwnd, LB_SETCURSEL, ItemIndex, 0)
  End If
End Sub
' Set Tab Stops (ListBox)
Public Sub SetTabsTops(ByVal LBHwnd As Long)
  Dim tabsets&(2)
  tabsets(0) = 45
  tabsets(1) = 110
  Call SendMessageLongByRef(LBHwnd, LB_SETTABSTOPS, 2, tabsets(0))
End Sub
' Increase Performance of Adding Data Into
' ComboBox and ListBox
Private Sub LockWinUpdate(ByVal hwndLock As Long)
  Call LockWindowUpdate(hwndLock)
End Sub
```

