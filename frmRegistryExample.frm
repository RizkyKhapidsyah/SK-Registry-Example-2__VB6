VERSION 5.00
Begin VB.Form frmRegistryExample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registry Example"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   Icon            =   "frmRegistryExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7800
      TabIndex        =   9
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdEnumValues 
      Caption         =   "Enum Reg Key Values"
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdEnumKey 
      Caption         =   "Enum Reg Keys"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdChangeKeyValue 
      Caption         =   "Change Reg Key Value"
      Height          =   495
      Left            =   7800
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdImportKey 
      Caption         =   "Import Reg Key"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdExportKey 
      Caption         =   "Export Reg Key"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdWriteKey 
      Caption         =   "Write Reg Key"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdReadKey 
      Caption         =   "Read Reg Key"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdDeleteKeyValue 
      Caption         =   "Delete Reg Key Value"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdDeleteKey 
      Caption         =   "Delete Reg Key"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmRegistryExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  frmRegistryExample
'  form used to demo the modRegistry
'=========================================================================================
'  Created By: Marc Cramer
'  Published Date: 04/18/2001
'  Copyright Date: 04/18/2001
'  WebSite: www.mkccomputers.com
'=========================================================================================
Option Explicit
'=========================================================================================
Private Sub cmdChangeKeyValue_Click()
' routine to overwrite current registry values
  On Error Resume Next
  modRegistry.WriteRegKey REG_SZ, HKEY_CURRENT_USER, "modRegistryTest\Directory1", "String Test", "0"
  modRegistry.WriteRegKey REG_BINARY, HKEY_CURRENT_USER, "modRegistryTest\Directory1", "Binary Test", "0"
  modRegistry.WriteRegKey REG_DWORD, HKEY_CURRENT_USER, "modRegistryTest\Directory1", "DWORD Test", "0"
  
  modRegistry.WriteRegKey REG_SZ, HKEY_CURRENT_USER, "modRegistryTest\Directory2", "String Test", "0"
  modRegistry.WriteRegKey REG_BINARY, HKEY_CURRENT_USER, "modRegistryTest\Directory2", "Binary Test", "0"
  modRegistry.WriteRegKey REG_DWORD, HKEY_CURRENT_USER, "modRegistryTest\Directory2", "DWORD Test", "0"
  
  modRegistry.WriteRegKey REG_SZ, HKEY_CURRENT_USER, "modRegistryTest\Directory3", "String Test", "0"
  modRegistry.WriteRegKey REG_BINARY, HKEY_CURRENT_USER, "modRegistryTest\Directory3", "Binary Test", "0"
  modRegistry.WriteRegKey REG_DWORD, HKEY_CURRENT_USER, "modRegistryTest\Directory3", "DWORD Test", "0"
  
  MsgBox "Writting to the Registry Complete", vbInformation, "Registry Operation Complete"
End Sub 'cmdChangeKeyValue_Click()
'=========================================================================================
Private Sub cmdDeleteKey_Click()
' routine to delete registry key
  On Error Resume Next
  modRegistry.DeleteRegKey HKEY_CURRENT_USER, "modRegistryTest", "Directory1"
  modRegistry.DeleteRegKey HKEY_CURRENT_USER, "modRegistryTest", "Directory2"
  modRegistry.DeleteRegKey HKEY_CURRENT_USER, "modRegistryTest", "Directory3"
  modRegistry.DeleteRegKey HKEY_CURRENT_USER, "modRegistryTest", ""
  
  MsgBox "Deleting Registry Keys Complete", vbInformation, "Registry Operation Complete"
End Sub 'cmdDeleteKey_Click()
'=========================================================================================
Private Sub cmdDeleteKeyValue_Click()
' routine to delete registry key values
  On Error Resume Next
  modRegistry.DeleteRegKeyValue HKEY_CURRENT_USER, "modRegistryTest\Directory1", "String Test"
  modRegistry.DeleteRegKeyValue HKEY_CURRENT_USER, "modRegistryTest\Directory1", "Binary Test"
  modRegistry.DeleteRegKeyValue HKEY_CURRENT_USER, "modRegistryTest\Directory1", "DWORD Test"
  
  modRegistry.DeleteRegKeyValue HKEY_CURRENT_USER, "modRegistryTest\Directory2", "String Test"
  modRegistry.DeleteRegKeyValue HKEY_CURRENT_USER, "modRegistryTest\Directory2", "Binary Test"
  modRegistry.DeleteRegKeyValue HKEY_CURRENT_USER, "modRegistryTest\Directory2", "DWORD Test"
  
  modRegistry.DeleteRegKeyValue HKEY_CURRENT_USER, "modRegistryTest\Directory3", "String Test"
  modRegistry.DeleteRegKeyValue HKEY_CURRENT_USER, "modRegistryTest\Directory3", "Binary Test"
  modRegistry.DeleteRegKeyValue HKEY_CURRENT_USER, "modRegistryTest\Directory3", "DWORD Test"
  
  MsgBox "Deleting Registry Key Values Complete", vbInformation, "Registry Operation Complete"
End Sub 'cmdDeleteKeyValue_Click()
'=========================================================================================
Private Sub cmdEnumKey_Click()
' routine to enumerate all registry keys
  On Error Resume Next
  Dim Message As String
  Dim NewMessage() As String
  Dim Counter As Integer

  Message = modRegistry.EnumerateRegKeys(HKEY_CURRENT_USER, "modRegistryTest")

  NewMessage = Split(Message, ",")
  Message = ""
  For Counter = LBound(NewMessage) To UBound(NewMessage)
    Message = Message & NewMessage(Counter) & vbCrLf
  Next Counter

  MsgBox Message, vbInformation, "Registry Operation Complete"
End Sub 'cmdEnumKey_Click()
'=========================================================================================
Private Sub cmdEnumValues_Click()
' routine to enumerate all registry keys values
  On Error Resume Next
  Dim Message As String
  Dim NewMessage() As String
  Dim Counter As Integer

  Message = modRegistry.EnumerateRegKeyValues(HKEY_CURRENT_USER, "modRegistryTest\Directory1")
  Message = Message & vbCrLf & "," & modRegistry.EnumerateRegKeyValues(HKEY_CURRENT_USER, "modRegistryTest\Directory2")
  Message = Message & vbCrLf & "," & modRegistry.EnumerateRegKeyValues(HKEY_CURRENT_USER, "modRegistryTest\Directory3")

  NewMessage = Split(Message, ",")
  Message = ""
  For Counter = LBound(NewMessage) To UBound(NewMessage)
    Message = Message & NewMessage(Counter) & vbCrLf
  Next Counter
  
  NewMessage = Split(Message, "*")
  Message = ""
  For Counter = LBound(NewMessage) To UBound(NewMessage)
    Message = Message & NewMessage(Counter) & vbCrLf
  Next Counter

  MsgBox Message, vbInformation, "Registry Operation Complete"
End Sub 'cmdEnumValues_Click()
'=========================================================================================
Private Sub cmdExit_Click()
' quitting time...
  On Error Resume Next
  Unload Me
  End
End Sub 'cmdExit_Click()
'=========================================================================================
Private Sub cmdExportKey_Click()
' routine to export registry key
  On Error Resume Next
  
  modRegistry.ExportRegKey HKEY_CURRENT_USER, "modRegistryTest", App.Path & "\TestExport.txt"
  
  MsgBox "Exporting the Registry Key Complete", vbInformation, "Registry Operation Complete"
End Sub 'cmdExportKey_Click()
'=========================================================================================
Private Sub cmdImportKey_Click()
' routine to import and overwrite current registry key
  On Error Resume Next
  
  modRegistry.ImportRegKey HKEY_CURRENT_USER, "modRegistryTest", App.Path & "\TestExport.txt"
  
  MsgBox "Importing the Registry Key Complete", vbInformation, "Registry Operation Complete"
End Sub 'cmdImportKey_Click()
'=========================================================================================
Private Sub cmdReadKey_Click()
' routine to read a registry key
  On Error Resume Next
  Dim Message As String

  Message = Message & "modRegistryTest\Directory1\String Test: " & modRegistry.ReadRegKey(HKEY_CURRENT_USER, "modRegistryTest\Directory1", "String Test", "NO KEY FOUND") & vbCrLf
  Message = Message & "modRegistryTest\Directory1\Binary Test: " & modRegistry.ReadRegKey(HKEY_CURRENT_USER, "modRegistryTest\Directory1", "Binary Test", "NO KEY FOUND") & vbCrLf
  Message = Message & "modRegistryTest\Directory1\DWORD Test: " & modRegistry.ReadRegKey(HKEY_CURRENT_USER, "modRegistryTest\Directory1", "DWORD Test", "NO KEY FOUND") & vbCrLf
  
  Message = Message & vbCrLf
  
  Message = Message & "modRegistryTest\Directory2\String Test: " & modRegistry.ReadRegKey(HKEY_CURRENT_USER, "modRegistryTest\Directory2", "String Test", "NO KEY FOUND") & vbCrLf
  Message = Message & "modRegistryTest\Directory2\Binary Test: " & modRegistry.ReadRegKey(HKEY_CURRENT_USER, "modRegistryTest\Directory2", "Binary Test", "NO KEY FOUND") & vbCrLf
  Message = Message & "modRegistryTest\Directory2\DWORD Test: " & modRegistry.ReadRegKey(HKEY_CURRENT_USER, "modRegistryTest\Directory2", "DWORD Test", "NO KEY FOUND") & vbCrLf
  
  Message = Message & vbCrLf
  
  Message = Message & "modRegistryTest\Directory3\String Test: " & modRegistry.ReadRegKey(HKEY_CURRENT_USER, "modRegistryTest\Directory3", "String Test", "NO KEY FOUND") & vbCrLf
  Message = Message & "modRegistryTest\Directory3\Binary Test: " & modRegistry.ReadRegKey(HKEY_CURRENT_USER, "modRegistryTest\Directory3", "Binary Test", "NO KEY FOUND") & vbCrLf
  Message = Message & "modRegistryTest\Directory3\DWORD Test: " & modRegistry.ReadRegKey(HKEY_CURRENT_USER, "modRegistryTest\Directory3", "DWORD Test", "NO KEY FOUND") & vbCrLf

  MsgBox Message, vbInformation, "Registry Operation Complete"
End Sub 'cmdReadKey_Click()
'=========================================================================================
Private Sub cmdWriteKey_Click()
' routine to write a registry key
  On Error Resume Next
  modRegistry.WriteRegKey REG_SZ, HKEY_CURRENT_USER, "modRegistryTest\Directory1", "String Test", "Test String 1"
  modRegistry.WriteRegKey REG_BINARY, HKEY_CURRENT_USER, "modRegistryTest\Directory1", "Binary Test", "1"
  modRegistry.WriteRegKey REG_DWORD, HKEY_CURRENT_USER, "modRegistryTest\Directory1", "DWORD Test", "1"
  
  modRegistry.WriteRegKey REG_SZ, HKEY_CURRENT_USER, "modRegistryTest\Directory2", "String Test", "Test String 2"
  modRegistry.WriteRegKey REG_BINARY, HKEY_CURRENT_USER, "modRegistryTest\Directory2", "Binary Test", "2"
  modRegistry.WriteRegKey REG_DWORD, HKEY_CURRENT_USER, "modRegistryTest\Directory2", "DWORD Test", "2"
  
  modRegistry.WriteRegKey REG_SZ, HKEY_CURRENT_USER, "modRegistryTest\Directory3", "String Test", "Test String 3"
  modRegistry.WriteRegKey REG_BINARY, HKEY_CURRENT_USER, "modRegistryTest\Directory3", "Binary Test", "3"
  modRegistry.WriteRegKey REG_DWORD, HKEY_CURRENT_USER, "modRegistryTest\Directory3", "DWORD Test", "3"

  MsgBox "Writting to the Registry Complete", vbInformation, "Registry Operation Complete"
End Sub 'cmdWriteKey_Click()
'=========================================================================================
