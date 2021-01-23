# vb6jsonx
Vb6jsonx is a COM component that extends VBA-JSON <https://github.com/vba-tools/vba-json>. It increases the simplicity and usability of **creating** or **querying** JSON objects.

## Installing
First register vb6jsonx.dll with  regsvr32.exe  , and then reference it in your project.
## Examples
```vb
Public Function Creating() As String
    Dim i As Long
    Dim Request As New JsonObject

    With Request.ReNew()
        .AddString "name", "bearx"
        .AddNumber "weight", 61
        .AddBoolean "sex", True
        .AddNull "xxx"

        With .NewObject("subobject")
            .AddString "item1", "a"
            .AddNumber "item2", 123.456
        End With

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

Public Sub Querying()
    Dim Request As New JsonObject
    Dim queryArray As New JsonArray
    Dim queryObject As New JsonObject
    
    Dim jsonStr As String
    
    jsonStr = Creating()
    
    Request.OfJSON jsonStr
    
    Debug.Print Request.Query("name")
    Debug.Print Request.Query("subobject.item1")
    Debug.Print Request.Query("subobject.item2")
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
