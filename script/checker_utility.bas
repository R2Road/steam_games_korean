REM  *****  BASIC  *****



Function IsMultiByte( b() as Byte )

	'
	' 유니코드 범위 체크 : English + Latin : https://en.wikipedia.org/wiki/List_of_Unicode_characters
	'
	IsMultiByte = ( b( 1 ) <> 0 )

End Function



Function IsNumber( b() as Byte )


	'
	' Multibyte Check
	'
	If b( 1 ) <> 0 Then
		IsNumber = False
	Else
		
		'
		' 유니코드 범위 체크 : 숫자 : 0030 ~ 0039
		'
		IsNumber = ( b( 0 ) >= &H30 And b( 0 ) <= &H39 )
		
	End If


End Function



Function IsKorean( b() as Byte )

	
	'
	' Multibyte Check
	'
	If b( 1 ) = 0 Then
		IsKorean = False
	Else
		'
		' 유니코드 범위 체크 : 한글 : AC00 ~ D7FF
		'
		IsKorean = ( b( 1 ) >= &HAC And b( 1 ) < &HD8 )
	End If
	

End Function



Function IsEnglish( b() as Byte )


	'
	' Multibyte Check
	'
	If b( 1 ) <> 0 Then
		IsEnglish = False
	Else
		
		'
		' 유니코드 범위 체크
		' 대문자 : 0041 ~ 005A
		' 소문자 : 0061 ~ 007A
		'
		IsEnglish = ( b( 0 ) >= &H41 And b( 0 ) <= &H5A ) Or ( b( 0 ) >= &H61 And b( 0 ) <= &H7A )
		
	End If


End Function
