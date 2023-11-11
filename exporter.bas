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
' 한글 자모 목록
'
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



'
' 생성자
'
Private Sub Class_Initialize()

    MsgBox "Exporter : Initialize"
    
    '
	'
	'
	InitKoreanPartsList
    
End Sub

'
' 소멸자
'
Private Sub Class_Terminate()

    MsgBox "Exporter : Terminate"
    
End Sub ' Destructor



'
' 출력
'
Function ExportList2( sheet as Variant, active_area_h as Integer, key_index as Integer, sub_index as Integer, out_file as Variant )
	
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
		s = Left( sheet.getCellByPosition( key_index, i ).String, 1 )
		s = UCase( s )
		b = s
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

