REM  *****  BASIC  *****

Option Explicit


Type ActiveArea
	w as Integer
	h as Integer
End Type



Sub Main

	'
	'
	'
    Dim sheet as Object
    sheet = ThisComponent.Sheets.getByName( "list" )
    
    
    '
    ' 읽어야할 데이터 영역 가져오기
    '
    Dim active_area as ActiveArea
    active_area = CalculateSheetActiveArea( sheet )
    MsgBox( active_area.w & " : " & active_area.h )
    
End Sub



Function CalculateSheetActiveArea( sheet as Variant ) as ActiveArea
	
	Dim ret as ActiveArea
	
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
	'
	'
	CalculateSheetActiveArea = ret
	
End Function











