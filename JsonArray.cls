VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ======================================================================================
' Name:     JsonArray
' Author:   Bearx (597893979@qq.com)
' Date:     2021-01
'
' Copyright © 2020 Bearx for vb6jsonx
' --------------------------------------------------------------------------------------

Private Coll As New Collection

Public Function ReNew() As JsonArray
    Set Coll = New Collection
    Set ReNew = Me
End Function

Public Property Get NativeObject() As Collection
    Set NativeObject = Coll
End Property

Public Property Set NativeObject(ByVal value As Collection)
    Set Coll = value
End Property

Public Function NewObject() As JsonObject
    Dim obj As New JsonObject
    
    Coll.Add obj.NativeObject
    Set NewObject = obj
End Function

Public Function NewArray() As JsonArray
    Dim obj As New JsonArray
    
    Coll.Add obj.NativeObject
    Set NewArray = obj
End Function

Public Sub AddString(ByVal value As String)
    Coll.Add value
End Sub

Public Sub AddBoolean(ByVal value As Boolean)
    Coll.Add value
End Sub

Public Sub AddNumber(ByVal value As Variant)
    Coll.Add CDec(value)
End Sub

Public Function Query(ByVal path As String) As Variant
    Dim p As Variant
    Dim o As Variant
    Dim idx As Long
    Dim k As String
    Dim i As Long
    Dim n As Long
    
    If InStr(path, "/") Then
        p = Split(path, "/")
    ElseIf InStr(path, ".") Then
        p = Split(path, ".")
    Else
        p = Split(path, "")
    End If
    
    n = UBound(p)
    Set o = Coll
    
    For i = 0 To n
        p(i) = Trim(p(i))
        
        If p(i) = "" Then
            'continue
        ElseIf p(i) = "{COUNT}" Then
            o = o.Count
        ElseIf p(i) Like "(*)" Then
            If Not IsNumeric(p(i)) Then
                Err.Raise vbObjectError + 100, , "the string """ & p(i) & """ in current node is not numeric index at path:" & path
            Else
                idx = Abs(p(i))
                If TypeName(o) <> "Collection" Then
                    Err.Raise vbObjectError + 100, , "index (" & idx & ") in current node is not Collection at path:" & path
                ElseIf (idx > o.Count Or idx < 1) Then
                    Err.Raise vbObjectError + 100, , "index (" & idx & ") in current Collection is out of range at path:" & path
                ElseIf IsObject(o(idx)) Then
                    Set o = o(idx)
                Else
                    o = o(idx)
                End If
            End If
        Else
            k = p(i)
            If TypeName(o) <> "Dictionary" Then
                Err.Raise vbObjectError + 100, , "key """ & k & """ in current node is not Dictionary at path:" & path
            ElseIf Not o.Exists(k) Then
                Err.Raise vbObjectError + 100, , "key """ & k & """ in current Dictionary is not exists at path:" & path
            ElseIf IsObject(o(k)) Then
                Set o = o(k)
            Else
                o = o(k)
            End If
        End If
    Next i
    
    If IsObject(o) Then
        Set Query = o
    Else
        Query = o
    End If
End Function

Public Function ToJSON(Optional ByVal whoObject, Optional ByVal Whitespace As Variant) As String
    If IsMissing(whoObject) Then
        ToJSON = JsonConverter.ConvertToJson(Coll, Whitespace)
    Else
        ToJSON = JsonConverter.ConvertToJson(whoObject, Whitespace)
    End If
End Function

Public Sub OfJSON(ByVal JsonString As String)
    Set Coll = JsonConverter.ParseJson(JsonString)
End Sub


