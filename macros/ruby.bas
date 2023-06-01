REM  *****  BASIC  *****

REM RUBY
REM     TEXT - kanji or other text that requires annotation above it
REM     ANNOTATION - hiragana/katakana or other annotation

REM 	OUTPUT: html ruby element
Function RUBY(TEXT as String, ANNOTATION as String) as String
	If isEmpty(TEXT) Or Len(TEXT) = 0 Or TEXT = "0 results found." Then
		RUBY = ""
		Exit Function
	End If
	RUBY ="<ruby>" + TEXT + "<rt>" + ANNOTATION + "</rt></ruby>"
End Function

REM RB
REM     TEXT - Text which has ruby text in format <SEPARATOR>kanji<SEPARATOR>hiraganaOrKatakana<SEPARATOR>
REM     SEPARATOR - text <SEPARATOR> that should be used for parsing the String with ruby text passed to TEXT

REM 	OUTPUT: html ruby element
Function RB(TEXT as String, SEPARATOR as String) as String
	If isEmpty(TEXT) Or Len(TEXT) = 0 Then
		Exit Function
	End If
	regex = "\|[^\|]*\|[^\|]*\|"
	regex = replace(regex, "|", SEPARATOR)
	occs = GETOCCURENCES(regex, TEXT)
	For Each occ In occs
	    REM split by char
        splitOcc = Split(occ, SEPARATOR)
        kanjiStr = splitOcc(1)
        furiganaStr = splitOcc(2)
        rubyStr = RUBY(kanjiStr, furiganaStr)
        TEXT = replace(TEXT, occ, rubyStr)
	Next occ
	RB = TEXT
End Function

Private Function GETOCCURENCES(REGEX as String, FINDIN as String)  	
	Dim results(0) As String

    oTextSearch = CreateUnoService("com.sun.star.util.TextSearch")
    oOptions = CreateUnoStruct("com.sun.star.util.SearchOptions")
    oOptions.algorithmType = com.sun.star.util.SearchAlgorithms.REGEXP
    oOptions.searchString = REGEX
    oTextSearch.setOptions(oOptions)
    
    oFound = oTextSearch.searchForward(findIn, 0, Len(FINDIN))
    nmbOfResults = 0
    
    While oFound.subRegExpressions > 0
		nStart = oFound.startOffset(0)
        nEnd = oFound.endOffset(0)
        nmbOfResults = nmbOfResults + 1
        Redim Preserve results(nmbOfResults-1)
		res = Mid(FINDIN, nStart + 1, nEnd - nStart)
		results(nmbOfResults-1) = res
		oFound = oTextSearch.searchForward(FINDIN, nEnd, Len(FINDIN))
    Wend
	
    GETOCCURENCES = results
End Function
