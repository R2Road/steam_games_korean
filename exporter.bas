REM  *****  BASIC  *****

Option Explicit


Const StartX = 0
Const StartY = 2

Type Size
    w as Integer
    h as Integer
End Type



Function CalculateSheetActiveArea( sheet as Variant ) as Size
	
	Dim ret as Size
	Dim i, j, h as Integer
	
	'
	' W
	'
	j = sheet.Rows.Count - 1
	For i = StartX to j
		If sheet.getCellByPosition( i, 0 ).String = "" Then
			Exit For
		EndIf
	Next i
	ret.w = i - 1
		
	'
	' H
	'
	j = sheet.Rows.Count - 1
	For i = StartY to j
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



Function ExportList( sheet as Variant, active_area_h as Integer, key_index as Integer, sub_index as Integer, pf as Variant )
	
	Dim title, result as String
    Dim i, j as Integer
    For i = StartY to active_area_h
    
    	'
    	' Check Export Flag
    	'
    	If sheet.getCellByPosition( StartX, i ).String = "x" Then
    		GoTo Continue
    	EndIf
    	
    	
    	'
    	' Empty is End
    	'
    	If sheet.getCellByPosition( key_index, i ).String = "" Then
    		Exit For
    	EndIf
    	
    	
    	'
    	' Title : [ Key Title, Sub Title ] or [ Key Title ]
    	'
    	If sheet.getCellByPosition( sub_index, i ).String = "" Then
    		title = _
    				"[" _
	    		& 	sheet.getCellByPosition( key_index, i ).String _
    			&	"]"
    	Else
    		title = _
    				"[" _
	    		& 	sheet.getCellByPosition( key_index, i ).String _
	    		& 	" | " _
	    		& 	sheet.getCellByPosition( sub_index, i ).String _
    			&	"]"
    	EndIf
    	
    	
    	'
    	' Build Info
    	'
    	result = _
    			"####" _
    		& 	" " _
    		& 	title _
    		& 	"( " _
    		&  		sheet.getCellByPosition( 3, i ).String _
    		& 	" )" _
    		& 	" " _
    		&	"( " _
	    		& 	sheet.getCellByPosition( 4, i ).String _
	    		& 	" | " & sheet.getCellByPosition( 5, i ).String _
	    		& 	" | " & sheet.getCellByPosition( 6, i ).String _
	    		&	" | " & sheet.getCellByPosition( 7, i ).String _
    		& " )"
    	
    	
    	'
    	'
    	'
    	pf.WriteLine( result )
    	'MsgBox( result )
    	
    Continue:
    Next i
	
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
	' Sheet
	'
    Dim sheet as Object
    sheet = ThisComponent.Sheets.getByName( "list" )
    
    
    '
    ' Max X, Y
    '
    Dim active_area as Size
    active_area = CalculateSheetActiveArea( sheet )
    MsgBox( "Active Area : " & StartX & " : " & StartY & " ~ " & active_area.w & " : " & active_area.h )
    
    
    '
    ' Export List
    '
    On Error GoTo ERROR_END 'Error 발생시 File 해제 용도
    
    	'
    	'
    	'
    	Dim range4sort
    		range4sort = sheet.getCellRangeByPosition( StartX, StartY, active_area.w, active_area.h )
    	
    	Dim sort_field(1) as new com.sun.star.util.SortField
			sort_field(0).SortAscending = TRUE
		    sort_field(0).FieldType = com.sun.star.util.SortFieldType.ALPHANUMERIC
		    sort_field(1).SortAscending = TRUE
		    sort_field(1).FieldType = com.sun.star.util.SortFieldType.ALPHANUMERIC
	    
	    Dim sort_description(0) as new com.sun.star.beans.PropertyValue
		    sort_description(0).Name = "SortFields"
		    sort_description(0).Value = sort_field()
		    
    	
	    '
	    ' Write : Korean List
	    '
	    pf.WriteLine( "## 한국어 제목" & Chr( 10 ) )
		    sort_field(0).Field = 1
		    sort_field(1).Field = 2
		    sort_description(0).Value = sort_field()
		    range4sort.Sort( sort_description() )
	    ExportList( sheet, active_area.h, 1, 2, pf )
	    
	    
	    pf.WriteLine( Chr( 10 ) & Chr( 10 ) )
	    
	    
	    '
	    ' Write : Number, English List
	    '
	    pf.WriteLine( "## 숫자, 영어 제목" & Chr( 10 ) )
		    sort_field(0).Field = 2
		    sort_field(1).Field = 1
		    sort_description(0).Value = sort_field()
	    	range4sort.Sort( sort_description() )
	    ExportList( sheet, active_area.h, 2, 1, pf )
	    
	    
    	MsgBox( "Success" )
    	
    ERROR_END:
    
    
    '
    ' File Close
    '
    pf.CloseFile()
	pf = pf.Dispose()
    
End Sub











