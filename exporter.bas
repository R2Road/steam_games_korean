REM  *****  LibreOffice VBA  *****

Option Explicit


Const StartX = 0
Const StartY = 2

Type Size
    w as Integer
    h as Integer
End Type

Private list_initial_consonaant( 19 ) as String
Private list_vowel( 21 ) as String
Private list_final_consonaant( 28 ) as String

Function InitKoreanPartsList
	
	If list_initial_consonaant( 0 ) = "" Then
		
		list_initial_consonaant = Array( "ㄱ", "ㄲ", "ㄴ", "ㄷ", "ㄸ", "ㄹ", "ㅁ", "ㅂ", "ㅃ", "ㅅ", "ㅆ", "ㅇ" , "ㅈ", "ㅉ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ" )
		list_vowel = Array( "ㅏ", "ㅐ", "ㅑ", "ㅒ", "ㅓ", "ㅔ", "ㅕ", "ㅖ", "ㅗ", "ㅘ", "ㅙ", "ㅚ", "ㅛ", "ㅜ", "ㅝ", "ㅞ", "ㅟ", "ㅠ", "ㅡ", "ㅢ", "ㅣ" )
		list_final_consonaant = Array( "", "ㄱ", "ㄲ", "ㄳ", "ㄴ", "ㄵ", "ㄶ", "ㄷ", "ㄹ", "ㄺ", "ㄻ", "ㄼ", "ㄽ", "ㄾ", "ㄿ", "ㅀ", "ㅁ", "ㅂ", "ㅄ", "ㅅ", "ㅆ", "ㅇ", "ㅈ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ" )
		
	EndIf
	
End Function



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



Function ExportList( sheet as Variant, active_area_h as Integer, key_index as Integer, sub_index as Integer, out_file as Variant )
	
	'
	' 초성 분리및 출력용
	'
	Dim current_initial_consonant as Integer : current_initial_consonant = 0
	Dim last_initial_consonant as Integer : last_initial_consonant = -1
	Dim current_code as Long : current_code = -1
	Dim last_code as Long : last_code = -1
	Dim s as String
	Dim b() as Byte
	
	'
	'
	'
	Dim title, company, result as String
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
		' 초성 출력
		'
		s = sheet.getCellByPosition( key_index, i ).String
		b = Mid( s, 1, 1 )
		current_code = ConvertBytes2Code( b )
		
		If IsDecompositionEnable( current_code ) Then
			
			'
			' 한글
			'
			
			current_initial_consonant = Extract_InitialConsonant( current_code )
			
			If last_initial_consonant <> current_initial_consonant Then
				out_file.WriteLine( "#### " & list_initial_consonaant( Extract_InitialConsonant( current_code ) ) )
			End If
			
		Else
			
			'
			' 한글 이외의 문자는 code 비교만으로 다른 초성임을 확인 할 수 있다.
			'
			If last_code <> current_code Then
				out_file.WriteLine( "#### " & b )
			End If
		
		End If
		
		last_initial_consonant = current_initial_consonant
		last_code = current_code
    	
    	
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
		' Company : 개발사, 퍼블리셔 정보가 같다면 하나로 표시한다.
		'
		If sheet.getCellByPosition( 6, i ).String = sheet.getCellByPosition( 7, i ).String Then
			company = sheet.getCellByPosition( 6, i ).String
		Else
			company = sheet.getCellByPosition( 6, i ).String & " | " & sheet.getCellByPosition( 7, i ).String
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
				& 	" | " & company _
			& " )"
    	
    	
		'
		'
		'
		out_file.WriteLine( result )
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
	'
	'
	InitKoreanPartsList
	
	
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
		pf.WriteLine( "## 영어 제목" & Chr( 10 ) )
			sort_field(0).Field = 2
			sort_field(1).Field = 1
			sort_description(0).Value = sort_field()
			range4sort.Sort( sort_description() )
		ExportList( sheet, active_area.h, 2, 1, pf )
		
		
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
	' File Close
	'
	pf.CloseFile()
	pf = pf.Dispose()
    
End Sub











