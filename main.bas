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
	'
	'
	out_file.WriteLine( header_pf.ReadAll() )
	
	
	'
	' File Close
	'
	header_pf.CloseFile()
	header_pf = header_pf.Dispose()	

End Function



Function IsDecompositionEnable( code as Long )

	'
	' 분해 가능한 한글 범위 : AC00( 가 : 44032 ) ~ D7A3( 힣 : 55203 )
	'
	IsDecompositionEnable = ( code >= 44032 And code <= 55203 )

End Function
Function ConvertBytes2Code( b() as Byte ) as Long

	'
	' byte array를 하나의 수로 만든다.
	'
	
	Dim code as Long 'Integer : 16bit, Long : 32bit
	
	'
	' b( 1 )
	'
	code = b( 1 )
	code = code * 256 ' 256 : 2의 8승 : 왼쪽 shift 8
	
	'
	' b( 0 )
	'
	code = code + b( 0 )
	
	ConvertBytes2Code = code
		
End Function
Function Extract_InitialConsonant( code as Long ) as Long
	
	Dim result as Long
	
	'
	' 한글 결합식
	'
	' (초성 인덱스 * 21 + 중성 인덱스) * 28 + 종성 인덱스 + 0xAC00( 44032 : 가 )
	'
	
	'
	' 가 : 44032
	' 각 항목의 인덱스가 모두 0 일때 '가' 이다.
	'
	result = Int( code - 44032 ) '은근슬쩍 반올림을 하고 있어서 Int 를 사용 해서 정수부만 쓰도록 제한한다.
	
	'
	' 종성 떨구기
	'
	result = Int( result / 28 )
	
	'
	' 중성 떨구기
	'
	result = Int( result / 21 )
	
	Extract_InitialConsonant = result
	
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
	'
	'
	LoadFile( "header.txt", pf )
	
	
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
		
		
		pf.WriteLine( Chr( 10 ) & Chr( 10 ) )
		pf.WriteLine( "<br/><br/>" )
		
		
		'
		' Write : Korean List
		'
		pf.WriteLine( "## 한국어 제목" & Chr( 10 ) )
			sort_field(0).Field = 1
			sort_field(1).Field = 2
			sort_description(0).Value = sort_field()
			range4sort.Sort( sort_description() )
		obj_exporter.ExportList2( sheet, active_area.h, 1, 2, pf )
		
		
		pf.WriteLine( Chr( 10 ) & Chr( 10 ) )
		pf.WriteLine( "<br/><br/>" )
		
		
		'
		' Write : Number, English List
		'
		pf.WriteLine( "## 영어 제목" & Chr( 10 ) )
			sort_field(0).Field = 2
			sort_field(1).Field = 1
			sort_description(0).Value = sort_field()
			range4sort.Sort( sort_description() )
		obj_exporter.ExportList2( sheet, active_area.h, 2, 1, pf )
		
		
		MsgBox( "Success" )
		
		
		'
		' Rollback
		'
		sort_field(0).Field = 1
		sort_field(1).Field = 2
		sort_description(0).Value = sort_field()
		range4sort.Sort( sort_description() )
	
	ERROR_END:
	
	obj_exporter = Nothing
    
    
	'
	' File Close
	'
	pf.CloseFile()
	pf = pf.Dispose()
    
End Sub











