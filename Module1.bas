Attribute VB_Name = "Module1"
Option Explicit
'   *************************************************************************
'   *       This program designed by Troy M. Reed for test scanning a       *
'   *       hand held serial scanner.                                       *
'   *************************************************************************
'   *                         Design Date:  04/03/00                        *
'   *************************************************************************
'   * Module for Publics and declare functions                              *
'   *************************************************************************

'Public String
'
Public CheckMySettings
'
'Declare Function
'
Public Declare Function FlashWindow Lib "user32" _
  (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

