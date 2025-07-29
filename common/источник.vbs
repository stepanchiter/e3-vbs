Set e3Application = CreateObject( "CT.Application" ) 
 
Set job = e3Application.CreateJobObject()
 
Set symbol = job.CreateSymbolObject()
 
 
 
Dim symbolName : symbolName = "Источник_1"
 
Dim symbolVersion : symbolVersion = "1"
 
 
 
isLoaded = symbol.Load( symbolName, symbolVersion )        'load symbol from database
 
If isLoaded > 0 Then    
 
    
 
    result = symbol.PlaceInteractively()
 
    If result = 0 Then
 
        message = "Error placing " & symbolName & " version " & symbolVersion
 
    Else
 
        message = "Symbol " & result & " placed interactively"                
 
    End If
 
    e3Application.PutInfo 0, message        'output result of operation    
 
            
 
End If
 
 
 
Set symbol = Nothing
 
Set job = Nothing
 
Set e3Application = Nothing
 
