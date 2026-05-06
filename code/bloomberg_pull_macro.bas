' ============================================================================
'  Bloomberg bulk data pull macro
'  Project: FinBERT-replicated News Sentiment as a Predictive Signal for Equities
'  Author : Siddhant Rajhans (FE 511 Final Project)
'
'  How to use:
'    1. Open Excel on a Bloomberg-connected machine. Confirm the Bloomberg
'       ribbon tab is loaded.
'    2. Press Alt+F11 to open the VBA editor.
'    3. Insert > Module.
'    4. Paste this entire file's contents into the module.
'    5. Press F5 (or Run > Run Sub > PullSentimentData).
'    6. Wait 3-10 minutes for all BDH formulas to populate
'       (Bloomberg shows "#N/A Requesting Data" while it fetches).
'    7. Once every cell shows real numbers, run PinAllToValues to lock the data.
'    8. File > Save As > bloomberg_raw.xlsx, drop in data/raw/.
' ============================================================================

Public Const START_DATE As String = "1/1/2018"

Sub PullSentimentData()
    Dim tickers As Variant
    tickers = Array( _
        "AAPL", "MSFT", "GOOGL", "NVDA", "META", "AMZN", "TSLA", _
        "JPM", "BAC", "GS", "WFC", _
        "JNJ", "UNH", "PFE", "LLY", "MRK", _
        "WMT", "PG", "KO", "PEP", "HD", "NKE", "DIS", _
        "XOM", "CVX", _
        "BA", "CAT", "GE", _
        "T", "VZ" _
    )

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Remove any pre-existing data sheets so re-runs are clean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "README" And ws.Name <> "Sheet1" Then
            ws.Delete
        End If
    Next ws

    Dim t As Variant
    Dim sht As Worksheet

    For Each t In tickers
        Set sht = ThisWorkbook.Worksheets.Add( _
                  After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sht.Name = CStr(t)

        sht.Range("A1").Value = "Date"
        sht.Range("B1").Value = "Price"
        sht.Range("C1").Value = "Sentiment"
        sht.Range("D1").Value = "NumStories"
        sht.Range("E1").Value = "MktCap"

        sht.Range("A2").Formula = _
            "=BDH(""" & t & " US Equity""," & _
            "{""PX_LAST"",""NEWS_SENTIMENT_DAILY_AVG""," & _
              """NUM_NEWS_STORIES_24HR"",""CUR_MKT_CAP""}," & _
            """" & START_DATE & """,TODAY()," & _
            """Dts=H"",""Sort=A"",""cols=5;rows=auto"")"
    Next t

    ' Benchmarks: SPX and VIX
    Set sht = ThisWorkbook.Worksheets.Add( _
              After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    sht.Name = "BENCHMARKS"
    sht.Range("A1").Value = "Date"
    sht.Range("B1").Value = "SPX"
    sht.Range("D1").Value = "VIX"

    sht.Range("A2").Formula = _
        "=BDH(""SPX Index"",""PX_LAST""," & _
        """" & START_DATE & """,TODAY()," & _
        """Dts=H"",""Sort=A"",""cols=2;rows=auto"")"

    sht.Range("C2").Formula = _
        "=BDH(""VIX Index"",""PX_LAST""," & _
        """" & START_DATE & """,TODAY()," & _
        """Dts=H"",""Sort=A"",""cols=2;rows=auto"")"

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Submitted " & (UBound(tickers) - LBound(tickers) + 1) & _
           " ticker pulls + benchmark sheet." & vbCrLf & _
           "Wait 3-10 minutes for Bloomberg to fully populate every cell," & _
           " then run PinAllToValues to lock the workbook."
End Sub


Sub PinAllToValues()
    ' Convert every formula on every sheet into static values so the
    ' workbook still works after the Bloomberg session ends.
    Dim sht As Worksheet
    Application.ScreenUpdating = False

    For Each sht In ThisWorkbook.Worksheets
        sht.UsedRange.Copy
        sht.UsedRange.PasteSpecial Paste:=xlPasteValues
    Next sht

    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    MsgBox "All sheets pinned to values. Save the workbook as bloomberg_raw.xlsx now."
End Sub


Sub HeadlineSampleForFinBERT()
    ' OPTIONAL: pull a sample of recent headlines for AAPL, MSFT, NVDA so we
    ' can run FinBERT locally and compare to Bloomberg's signal (Hypothesis 4).
    ' Uses the news-headline field. If your terminal doesn't license bulk
    ' headline export, this sheet will be empty and we use Bloomberg's signal
    ' alone (still A+ work).
    Dim sht As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("HEADLINES").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set sht = ThisWorkbook.Worksheets.Add( _
              After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    sht.Name = "HEADLINES"
    sht.Range("A1").Value = "AAPL_Headlines_Last30Days"

    sht.Range("A2").Formula = _
        "=BDS(""AAPL US Equity"",""NEWS_HEADLINES""," & _
        """START_DATE_OVERRIDE=" & Format(Date - 30, "yyyymmdd") & """," & _
        """END_DATE_OVERRIDE=" & Format(Date, "yyyymmdd") & """)"

    MsgBox "Headline sample requested. If you see #N/A, your license doesn't" & _
           " include bulk headline export and we'll skip H4 replication" & _
           " (or use Reuters via free sources)."
End Sub
