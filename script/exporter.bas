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



Public target_laugnage as Long



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
	' 왠지 Enum 이 에러난다. 일단 숫자로 대충 때려박기
	'
	target_laugnage = 0
	
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
'
'
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



'
' 언어 타입 확인
'

Function IsTargetLanguage( b() as Byte ) as Boolean

	Select Case target_laugnage
		Case 1: 'Korean
			IsTargetLanguage = IsKorean( b )
			
		Case 2: 'Not Korean
			IsTargetLanguage = Not IsKorean( b )
			
		Case 3: 'English
			IsTargetLanguage = IsEnglish( b )
		
		Case 4: 'Not English
			IsTargetLanguage = Not IsEnglish( b )
			
		Case Else
			IsTargetLanguage = False
		
	End Select

End Function



'
' 출력
'
Function ExportList( sheet as Variant, active_area_start_y as Integer, active_area_end_y as Integer, key_index as Integer, sub_index as Integer, out_file as Variant )
	
	'
	' Enum이 Cell 가져올 때 문제가 있다.
	'
	Dim eSI_STEAM_LINK as Long : eSI_STEAM_LINK = 3
	Dim eSI_RELEASE_YEAR as Long : eSI_RELEASE_YEAR = 4
	Dim eSI_GENRE as Long : eSI_GENRE = 7
	
	Dim eSI_DEVELOPER_1 as Long : eSI_DEVELOPER_1 = 8
	Dim eSI_DEVELOPER_2 as Long : eSI_DEVELOPER_2 = 9
	Dim eSI_DEVELOPER_3 as Long : eSI_DEVELOPER_3 = 10
	
	Dim eSI_PUBLISHER_1 as Long : eSI_PUBLISHER_1 = 11
	Dim eSI_PUBLISHER_2 as Long : eSI_PUBLISHER_2 = 12
	Dim eSI_PUBLISHER_3 as Long : eSI_PUBLISHER_3 = 13
	
	
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
	Dim i as Integer
	For i = active_area_start_y to active_area_end_y
    
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
		' 첫 글자 잘라내기
		'
		s = Left( sheet.getCellByPosition( key_index, i ).String, 1 )
		s = UCase( s )
		b = s
		
		
		
		'
		' 목표 언어 종류 검사
		'
		If IsTargetLanguage( b ) = False Then
			GoTo Continue
		End If
		
		
		
		'
		'
		'
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
		' Company
		'  > 개발사, 퍼블리셔 정보가 각 각 1개 이고
		'  > 개발사, 퍼블리셔 정보가 같다면
		'  > 개발사 정보만 표시한다.
		'
		
		
		If sheet.getCellByPosition( eSI_DEVELOPER_1, i ).String = sheet.getCellByPosition( eSI_PUBLISHER_1, i ).String _
			And ( sheet.getCellByPosition( eSI_DEVELOPER_2, i ).String = "" ) _
			And ( sheet.getCellByPosition( eSI_DEVELOPER_3, i ).String = "" ) _
			And ( sheet.getCellByPosition( eSI_PUBLISHER_2, i ).String = "" ) _
			And ( sheet.getCellByPosition( eSI_PUBLISHER_3, i ).String = "" ) Then
			
			company = sheet.getCellByPosition( eSI_DEVELOPER_1, i ).String
			
		Else
		
			'
			' 개발자 정보 1
			'
			company = sheet.getCellByPosition( eSI_DEVELOPER_1, i ).String
			
			'
			' 개발자 정보 2
			'
			If sheet.getCellByPosition( eSI_DEVELOPER_2, i ).String <> "" Then
				company = _
						company _
					& 	", " _
					& 	sheet.getCellByPosition( eSI_DEVELOPER_2, i ).String
			End If
			
			'
			' 개발자 정보 3
			'
			If sheet.getCellByPosition( eSI_DEVELOPER_3, i ).String <> "" Then
				company = _
						company _
					& 	", " _
					& 	sheet.getCellByPosition( eSI_DEVELOPER_3, i ).String
			End If
			
			
			
			'
			' 퍼블리셔 정보 1
			'
			company = _
						company _
					& 	" | " _
					& 	sheet.getCellByPosition( eSI_PUBLISHER_1, i ).String
					
			'
			' 퍼블리셔 정보 2
			'
			If sheet.getCellByPosition( eSI_PUBLISHER_2, i ).String <> "" Then
				company = _
						company _
					& 	", " _
					& 	sheet.getCellByPosition( eSI_PUBLISHER_2, i ).String
			End If
			
			'
			' 퍼블리셔 정보 3
			'
			If sheet.getCellByPosition( eSI_PUBLISHER_3, i ).String <> "" Then
				company = _
						company _
					& 	", " _
					& 	sheet.getCellByPosition( eSI_PUBLISHER_3, i ).String
			End If
			
		EndIf
    	
    	
    	
		'
		' Build Info
		'
		result = _
				"####" _
			& 	" " _
			& 	title _
			& 	"( " _
			&  		sheet.getCellByPosition( eSI_STEAM_LINK, i ).String _
			& 	" )" _
			& 	" " _
			&	"( " _
				& 	        sheet.getCellByPosition( eSI_RELEASE_YEAR, i ).String _
				& 	" | " & sheet.getCellByPosition( eSI_GENRE, i ).String _
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

