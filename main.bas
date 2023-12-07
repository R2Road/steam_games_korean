REM  *****  LibreOffice VBA  *****

Option Explicit


Const StartX = 0
Const StartY = 2

Type Size
    w as Integer
    h as Integer
End Type



Function CalculateSheetActiveArea( sheet as Variant ) as Size
	
	Dim ret as Size
	
	'
	' W
	'
	Dim cur_x as Long
	Dim end_x as Long : end_x = sheet.Columns.Count - 1
	
	For cur_x = StartX to end_x
		If sheet.getCellByPosition( cur_x, 0 ).String = "" Then
			cur_x = cur_x - 1
			Exit For
		EndIf
	Next cur_x
	ret.w = cur_x
		
	'
	' H
	'
	Dim cur_y as Long
	Dim end_y as Long : end_y = sheet.Rows.Count - 1
	
	For cur_y = StartY to end_y
		If sheet.getCellByPosition( 0, cur_y ).String = "" Then
			cur_y = cur_y - 1
			Exit For
		EndIf
	Next cur_y
	ret.h = cur_y
	
	'
	' Return
	'
	CalculateSheetActiveArea = ret
	
End Function



Function LoadFile( header_fine_name as String, out_file as Variant )

	'
	' File Open
	'
	Dim file_path as String
	file_path = ( Tools.Strings.DirectoryNameoutofPath(ThisComponent.getURL(),"/" ) & "/" & header_fine_name )
	'MsgBox( file_path )
	
	Dim file_system As Variant
	file_system = CreateScriptService("FileSystem")
	
	Dim header_pf As Variant
	Set header_pf = file_system.OpenTextFile(file_path, file_system.ForReading)
	
	
	'
	' Write
	'
	out_file.WriteLine( header_pf.ReadAll() )
	
	
	'
	' File Close
	'
	header_pf.CloseFile()
	header_pf = header_pf.Dispose()	

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
	file_path = ( Tools.Strings.DirectoryNameoutofPath(ThisComponent.getURL(),"/" ) & "/" & "README.md" )
	MsgBox( file_path )
	
	Dim file_system As Variant
	file_system = CreateScriptService("FileSystem")
	
	Dim pf As Variant
	Set pf = file_system.CreateTextFile(file_path, Overwrite := true)
	
	
	
	'
	' Insert "Header.txt"
	'
	LoadFile( "header.txt", pf )
	
	
	'
	' Sheet "list"
	'
	Dim sheet as Object
	sheet = ThisComponent.Sheets.getByName( "list" )
    
    
	'
	' Max X, Y
	'
	Dim active_area as Size
	active_area = CalculateSheetActiveArea( sheet )
	MsgBox( _
							"Active Area : " _
		& 	Chr( 10 ) & 	"Start : "	& StartX 			& " : " 	& StartY _
		& 	Chr( 10 ) & 	"End  : " 	& active_area.w 	& " : " 	& active_area.h _
	)
	
	
	'
	' 게임 수 출력
	'
	pf.WriteLine( active_area.h - StartY & " 개" )
    
    
	'
	' Export List
	'
	Dim obj_exporter as Object
	Set obj_exporter = New exporter
	
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
		'
		'
		pf.WriteLine( Chr( 10 ) & Chr( 10 ) )
		pf.WriteLine( "<br/><br/>" )
		
		'
		' Write : Korean List
		'
			
			'
			' Sort
			'
			pf.WriteLine( "## 한국어 제목" & Chr( 10 ) )
				sort_field(0).Field = 1
				sort_field(1).Field = 2
				sort_description(0).Value = sort_field()
				range4sort.Sort( sort_description() )
				
			'
			' Write : Korean
			'
			obj_exporter.target_laugnage = 1
			obj_exporter.ExportList( sheet, active_area.h, 1, 2, pf )
			
			'
			' Write : Others
			'
			obj_exporter.target_laugnage = 2
			obj_exporter.ExportList( sheet, active_area.h, 1, 2, pf )
		
		
		
		
		'
		'
		'
		pf.WriteLine( Chr( 10 ) & Chr( 10 ) )
		pf.WriteLine( "<br/><br/>" )
		
		'
		' Write : English List
		'
		
			'
			' Sort
			'
			pf.WriteLine( "## 영어 제목" & Chr( 10 ) )
				sort_field(0).Field = 2
				sort_field(1).Field = 1
				sort_description(0).Value = sort_field()
				range4sort.Sort( sort_description() )
			
			'
			' Write : English
			'
			obj_exporter.target_laugnage = 3
			obj_exporter.ExportList( sheet, active_area.h, 2, 1, pf )
			
			'
			' Write : Others
			'
			obj_exporter.target_laugnage = 4
			obj_exporter.ExportList( sheet, active_area.h, 2, 1, pf )
		
		
		
		
		'
		'
		'
		MsgBox( "Success" )
		
		
		
		
		'
		' Rollback
		'
		sort_field(0).Field = 1
		sort_field(1).Field = 2
		sort_description(0).Value = sort_field()
		range4sort.Sort( sort_description() )
	
	ERROR_END:
	
	'
	'
	'
	obj_exporter = Nothing
    
    
	'
	' File Close
	'
	pf.CloseFile()
	pf = pf.Dispose()
    
End Sub











