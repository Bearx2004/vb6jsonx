# vb6jsonx
Vb6jsonx is a COM component that extends VB JSON. It increases the simplicity and usability of **creating** or **querying** JSON objects.
## installing
First register with Regsvr32.exe  vb6jsonx.dll , and then reference it in the project.
## examples
```vb
Public Function Creating() As String
    Dim i As Long
    Dim Request As New vb6jsonx.JsonObject

    With Request.ReNew()
        .AddString "name", "bearx"
        .AddNumber "weight", 61
        .AddBoolean "sex", True
        .AddNull "xxx"
        
        With .NewArray("love")
            .AddString "music"
            .AddString "painting"
        End With
        
        With .NewArray("detail")
            For i = 1 To 5
                With .NewObject()
                    .AddNumber "ID", i
                    .AddString "memo", "中国成都" & i
                    .AddString "create-date", Now
                    .AddNull "xxx"
                End With
            Next i
        End With
    End With
        
    Creating = Request.ToJSON()
    
End Function

Public Sub QueryJson()
    Dim Request As New vb6jsonx.JsonObject
    Dim queryArray As New JsonArray
    Dim queryObject As New JsonObject
    
    Dim jsonStr As String
    
    jsonStr = Creating()
    
    Request.OfJSON jsonStr
    
    Debug.Print Request.Query("name")
    Debug.Print Request.Query("detail.{COUNT}")
    Debug.Print Request.Query("detail.(1).memo")
    
    Set queryArray.NativeObject = Request.Query("detail")
    Debug.Print queryArray.Query("(1).memo")
    Set queryObject.NativeObject = queryArray.Query("(1)")
    Debug.Print queryObject.Query("memo")
    
    Debug.Print Request.ToJSON(Request.Query("detail"))
    Debug.Print Request.ToJSON(Request.Query("detail.(1)"))
    Debug.Print Request.ToJSON(, 2)
        
    Debug.Print Request.ToUrlEncoder()
End Sub
```
