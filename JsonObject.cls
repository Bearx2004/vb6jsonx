VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ======================================================================================
' Name:     JsonObject
' Author:   Bearx (597893979@qq.com)
' Date:     2021-01
'
' Copyright © 2020 Bearx for vb6jsonx
' --------------------------------------------------------------------------------------

Private Dict As New Dictionary

Public Function ReNew() As JsonObject
    Set Dict = New Dictionary
    Set ReNew = Me
End Function

Public Property Set NativeObject(ByVal value As Dictionary)
    Set Dict = value
End Property

Public Property Get NativeObject() As Dictionary
    Set NativeObject = Dict
End Property

Public Function NewObject(ByVal key As String) As JsonObject
    Dim obj As New JsonObject
    
    Dict.Add key, obj.NativeObject
    Set NewObject = obj
End Function

Public Function NewArray(ByVal key As String) As JsonArray
    Dim obj As New JsonArray
    
    Dict.Add key, obj.NativeObject
    Set NewArray = obj
End Function

Public Sub AddString(ByVal key As String, ByVal value As String)
    Dict.Add key, value
End Sub

Public Sub AddBoolean(ByVal key As String, ByVal value As Boolean)
    Dict.Add key, value
End Sub

Public Sub AddNumber(ByVal key As String, ByVal value As Variant)
    Dict.Add key, CDec(value)
End Sub

Public Sub AddNull(ByVal key As String)
    Dict.Add key, Null
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
    Set o = Dict
    
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
        ToJSON = JsonConverter.ConvertToJson(Dict, Whitespace)
    Else
        ToJSON = JsonConverter.ConvertToJson(whoObject, Whitespace)
    End If
End Function

Public Sub OfJSON(ByVal JsonString As String)
    Set Dict = JsonConverter.ParseJson(JsonString)
End Sub

Public Function ToUrlEncoder(Optional ByVal linkChar As String = "=", Optional ByVal linkItemChar As String = "&") As String
    Dim keys
    Dim values() As String
    Dim i As Long
    Dim n As Long
    
    keys = Dict.keys
    n = UBound(keys)
    If n < 0 Then
        Exit Function
    End If
    
    ReDim values(n)
    
    For i = 0 To n
        If IsObject(Dict(keys(i))) Then
            values(i) = keys(i) & linkChar & ToJSON(Dict(keys(i)))
        Else
            values(i) = keys(i) & linkChar & Dict(keys(i))
        End If
    Next i
    
    ToUrlEncoder = Join(values, linkItemChar)
End Function


