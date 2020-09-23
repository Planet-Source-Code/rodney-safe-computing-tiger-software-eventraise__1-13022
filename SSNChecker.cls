VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SSNChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event BadSsn(ByVal bad_ssn As String)

' Verify that the string the format 111-11-1111.
' If it does not, raise the BadSsn event.
Public Sub CheckSsn(ByVal ssn As String)
    If Not (ssn Like "###-##-####") Then
        RaiseEvent BadSsn(ssn)
    End If
End Sub
