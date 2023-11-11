REM  *****  BASIC  *****

'
' 
'
Option Compatible
Option ClassModule

'
' Enum 을 사용하려면 이 "Option VBASupport 1" 이 필요하다
'
Option VBASupport 1



'
' 생성자
'
Private Sub Class_Initialize()

    MsgBox "Exporter : Initialize"
    
End Sub

'
' 소멸자
'
Private Sub Class_Terminate()

    MsgBox "Exporter : Terminate"
    
End Sub ' Destructor
