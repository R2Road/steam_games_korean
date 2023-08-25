REM  *****  BASIC  *****

Option Explicit

Type Point
	x as Integer
	y as Integer
End Type

Type Size
	w as Integer
	h as Integer
End Type



Const StartX = 1
Const StartY = 2



Function CalculateSheetActiveArea( sheet as Variant ) as ActiveArea
	
	Dim ret as Size
	
	Dim i, j, h as Integer ' h 를 넣으면 두번째 j 설정 코드에서 문제가 없다. 이거 뭐야?
	
	
	'
	' W
	'
	j = sheet.Columns.Count - 1
	For i = 1 to j
		If sheet.getCellByPosition( i, 0 ).String = "" Then
			ret.w = i - 1
			Exit For
		EndIf
	Next i	
	
	
	'
	' H
	'
	j = sheet.Rows.Count - 1
	For i = 2 to j
		If sheet.getCellByPosition( 0, i ).String = "" Then
			Exit For
		EndIf
	Next i	
	
	ret.h = i - 1
	
	
	'
	' Return
	'
	CalculateSheetActiveArea = ret
	
End Function



Sub Main

	'
	'
	'
	GlobalScope.BasicLibraries.LoadLibrary("Tools") ' for Tools
	GlobalScope.BasicLibraries.LoadLibrary("ScriptForge") ' for FileSystem
	
	
	'
	' File Open
	'
	Dim file_path as String
	file_path = ( Tools.Strings.DirectoryNameoutofPath(ThisComponent.getURL(),"/" ) & "/" & "rme.txt" )
	MsgBox( file_path )
	
	Dim file_system As Variant
	file_system = CreateScriptService("FileSystem")
	
	Dim pf As Variant
	Set pf = file_system.CreateTextFile(file_path, Overwrite := true)
	
	
	'
	'
	'
    Dim sheet as Object
    sheet = ThisComponent.Sheets.getByName( "list" )
    
    
    '
    ' Max X, Y
    '
    Dim active_area as Size
    active_area = CalculateSheetActiveArea( sheet )
    MsgBox( active_area.w & " : " & active_area.h )
    
    
    '
    ' Write
    '
    Dim s as String
    Dim i, j as Integer
    For i = StartY to 2 'active_area.h
    
    	s = ""
    
    	
    	s = _
    			"####" _
    		& 	" " _
    		& 	"[" & sheet.getCellByPosition( 1, i ).String & " | " & sheet.getCellByPosition( 2, i ).String & "]( " &  sheet.getCellByPosition( 3, i ).String & " )" _
    		& 	" " _
    		&	"( " _
	    		& 	sheet.getCellByPosition( 4, i ).String _
	    		& 	" | " & sheet.getCellByPosition( 5, i ).String _
	    		& 	" | " & sheet.getCellByPosition( 6, i ).String _
	    		&	" | " & sheet.getCellByPosition( 7, i ).String _
    		& " )"
    	
    	pf.WriteLine( s )
    	'MsgBox( s )
    	
    Next i
    
    
    '
    ' File Close
    '
    pf.CloseFile()
	pf = pf.Dispose()
    
End Sub











