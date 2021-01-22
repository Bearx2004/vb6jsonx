# vb6jsonx
Vb6jsonx is a COM component that extends VB JSON. It increases the simplicity and usability of **creating** or **querying** JSON objects.
## installing
First register with Regsvr32.exe  vb6jsonx.dll , and then reference it in the project.
## examples
```vb
Public Function CreateJson() As String
    Dim i As Long
    Dim Request As New vb6jsonx.JsonObject

    With Request.ReNew()
        .AddString "name", "bearx"
        .AddNumber "age", 50
        .AddBoolean "sex", True
        .AddNull "xxx"
        With .NewArray("detail")
            For i = 1 To 5
                With .NewObject()
                    .AddNumber "ID", i
                    .AddString "memo", "bearx" & i
                    .AddString "create-date", Now
                End With
            Next i
        End With
    End With
        
    CreateJson = Request.ToJSON()
    
End Function

Public Sub QueryJson()
    Dim Request As New vb6jsonx.JsonObject
    Dim jsonStr As String
    
    jsonStr = CreateJson()
    
    Call Request.OfJSON(jsonStr)
    
    Debug.Print Request.Query("name")
    Debug.Print Request.Query("detail.{COUNT}")
    Debug.Print Request.Query("detail.(1).memo")
    
    Debug.Print Request.ToJSON(Request.Query("detail"))
    Debug.Print Request.ToJSON(Request.Query("detail.(1)"))
    Debug.Print Request.ToJSON()
    Debug.Print Request.ToJSON(,2)
    Debug.Print Request.ToUrlEncoder()
End Sub
```
