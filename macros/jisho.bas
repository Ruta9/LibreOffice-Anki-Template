REM  *****  BASIC  *****

REM **************** CONSTANTS ***********************

REM Value: ROMAJI_COLUMN 
REM        indicates which column's value will be sent to Jisho. The dropdown will be created in column ROMAJI_COLUMN - 1
Const ROMAJI_COLUMN = 4 REM <-- 4 is column D. 

REM Value: JISHO_RESULT_LIMIT 
REM        indicates how many top results will be shown in the dropdown
Const JISHO_RESULT_LIMIT = 5

REM Value: JISHO_API_URL 
REM        Jishio api url with keyword parameter
Const JISHO_API_URL = "https://jisho.org/api/v1/search/words?keyword="



REM **************** PUBLIC FUNCTIONS ***********************

REM CREATEDROPDOWNONCHANGE
REM     Sub used for sheet event binding. It checks whether the change occured in configured ROMAJI_COLUMN and if so,
REM		generates a dropdown to the left of the edited cell.
Sub CREATEDROPDOWNONCHANGE(oEvent As Variant)

	If (Not oEvent.supportsService("com.sun.star.sheet.SheetCell")) Or (Not oEvent.supportsService("com.sun.star.table.Cell")) Or (Not oEvent.supportsService("com.sun.star.sheet.SheetCellRange")) Or (Not oEvent.supportsService("com.sun.star.table.CellRange")) Then
		Exit Sub
	End If
	
	column = oEvent.CellAddress.Column
	row = oEvent.CellAddress.Row
	
	If column <> ROMAJI_COLUMN Then
		Exit Sub
	End If
	
	oSheet = thiscomponent.sheets.getByIndex(0)
	
	REM Get keyword
	oCell = oSheet.getCellByPosition(column, row)
	keyword = oCell.getString()
	limit = JISHO_RESULT_LIMIT
	values = CALLDICT(keyword, limit)
	
	REM Create a dropdown to the next cell to the right
	CREATEDROPDOWN(values, column-1, row)
End Sub

REM CREATEDROPDOWN
REM     VALUES - Jisho response
REM     COLUMN - column where dropdown should be placed
REM     ROW - row where dropdown should be placed
Sub CREATEDROPDOWN(VALUES as String, COLUMN as Integer, ROW as Integer)
	Dim default as String
	Dim oCell As Object
	
	default = GETDEFAULT(VALUES)

	oSheet = thiscomponent.sheets.getByIndex(0)
	oCell = oSheet.getCellByPosition(COLUMN, ROW)
	Validation = oCell.Validation
	Validation.Type = com.sun.star.sheet.ValidationType.LIST
	Validation.setOperator(com.sun.star.sheet.ConditionOperator.FORMULA)
	Validation.setFormula1("SPLITVALUES(""" + values + """)")
	oCell.Validation = Validation
	oCell.setString(default)
End Sub

REM KANJI
REM     TEXT - dropdown's selected value

REM 	OUTPUT: kanji from dropdown's selected value
Function KANJI(TEXT as String) As String
	If isEmpty(TEXT) Or Len(TEXT) = 0 Then
		KANJI = ""
		Exit Function
	End If
	kanjiArr = Split(TEXT, "  |   ")
	If isArray(kanjiArr) And UBound(kanjiArr) >= 1 Then
		KANJI = kanjiArr(0)
	Else
		KANJI = ""
	End If
End Function

REM READING
REM     TEXT - dropdown's selected value

REM 	OUTPUT: hiragana/katakana from dropdown's selected value
Function READING(TEXT as String) As String
	If isEmpty(TEXT) Or Len(TEXT) = 0 Then
		READING = ""
		Exit Function
	End If
	readingArr = Split(TEXT, "   |   ")
	If isArray(readingArr) And UBound(readingArr) >= 1 Then
		READING = readingArr(1)
	Else
		READING = ""
	End If
End Function


REM **************** PRIVATE FUNCTIONS ***********************


Private Function GETELEMENT(JSON as String, ROW as Integer) as String
	Dim pattern as String
	Dim result as String
	pattern = "data.get(" + ROW + ")"

   	oService = createUNOService("com.github.binnarywolf.LibreOfficeGetRestPlugin")
   	
	result = oService.PARSEJSON(JSON, pattern)
	GETELEMENT = Replace(result,"No such node found.","")
End Function

Private Function GETKANJI(ELEMENT as String) As String
	Dim pattern as String
	Dim result as String
	pattern = "japanese.get(0).word"

   	oService = createUNOService("com.github.binnarywolf.LibreOfficeGetRestPlugin")
   	
	result = oService.PARSEJSON(ELEMENT, pattern)
	GETKANJI = Replace(result,"No such node found.","")
End Function

Private Function GETREADING(ELEMENT as String) As String
	Dim pattern as String
	Dim result as String
	pattern = "japanese.get(0).reading"

   	oService = createUNOService("com.github.binnarywolf.LibreOfficeGetRestPlugin")
   	
	result = oService.PARSEJSON(ELEMENT, pattern)
	GETREADING = Replace(result,"No such node found.","")
End Function

Private Function GETTRANSLATIONS(ELEMENT as String) As String
	Dim definitionPattern as String
	Dim definitionJson as String
	Dim count as Integer
	Dim iCount as Integer
	Dim definitions as String
	Dim result as String
	
	oService = createUNOService("com.github.binnarywolf.LibreOfficeGetRestPlugin")
	
	pattern = "senses"
	definitions = oService.PARSEJSON(ELEMENT, pattern)
	count = OCCURENCES(definitions, "english_definitions")
	result = ""
    For iCount = 0 To count - 1
        pattern = "get(" + iCount + ").english_definitions"
        result = result + oService.PARSEJSON(definitions, pattern) + " "
    Next iCount
	
	GETTRANSLATIONS = result
End Function


Private Function GETDROPDOWNLINE(JSON as String, ROW as Integer) As String
	Dim line as String
	Dim values(2) as String
	Dim kanji as String
	Dim reading as String
	Dim translations as String
	Dim index as Integer
	
	line = GETELEMENT(JSON, ROW)
	kanji = GETKANJI(line)
	reading = GETREADING(line)
	translations = GETTRANSLATIONS(line)
		
	GETDROPDOWNLINE = Join(Array(kanji, reading, translations), "   |   ")
End Function

Private Function OCCURENCES(TEXT as String, PHRASE as String) As String
	OCCURENCES = (Len(TEXT) - Len(Replace(TEXT, PHRASE, "")))/Len(PHRASE)
End Function

Private Function CALLDICT(KEYWORD as String, LIMIT as Integer) As String
	Dim Json as String
	Dim NmbOfResults as Integer
	Dim count as Integer
	Dim iCount as Integer
	Dim elements(LIMIT-1) as String
	Dim result as String
	
	If isEmpty(KEYWORD) Or Len(KEYWORD) = 0 Then
		CALLDICT = ""
		Exit Function
	End If
	
	oService = createUNOService("com.github.binnarywolf.LibreOfficeGetRestPlugin")
	
	Json = oService.GET(JISHO_API_URL + KEYWORD)
	NmbOfResults = OCCURENCES(Json, "slug")

    If NmbOfResults = 0 Then
        CALLDICT = "0 results found."
        Exit Function
    End If
    
    If NmbOfResults < LIMIT Then 
    	count = NmbOfResults 
    Else 
    	count = LIMIT
    End If
    
	result = ""
	
	For iCount = 0 To count - 1
        elements(iCount) = GETDROPDOWNLINE(Json, iCount)
    Next iCount
    
    CALLDICT = replace(Join(elements,";"), Chr(34), "")
End Function

Private Function GETDEFAULT(TEXT as String) As String
	If isEmpty(TEXT) Or Len(TEXT) = 0 Or TEXT = "0 results found." Then
		GETDEFAULT = ""
		Exit Function
	End If
	GETDEFAULT = Split(TEXT,";")(0)
End Function

Private Function SPLITVALUES(TEXT as String)
	If isEmpty(TEXT) Or Len(TEXT) = 0 Or TEXT = "0 results found." Then
		SPLITVALUES = ""
		Exit Function
	End If
	SPLITVALUES = Split(TEXT,";")
End Function
