' Written by Ross Sharma
' April 30, 2015
' https://github.com/ross-sharma/excel-vba-settingsreader
'
' ross.sharma86@gmail.com
' www.elance.com/s/rosssharma
'
' v1.0

Option Explicit


Private pWorksheet As Worksheet
Private pUseTestValues As Boolean
Private pCache As Object 'Scripting.Dictionary
Private pCacheEnabled As Boolean
Private pCaseSensitive As Boolean
Private pTestSettingPrefix As String

Private Sub Class_Initialize()
    Const DEFAULT_TEST_SETTING_PREFIX As String = "_test_"
    
    pUseTestValues = False
    Set pCache = CreateObject("Scripting.Dictionary")
    pCacheEnabled = False
    pCaseSensitive = False
    pTestSettingPrefix = DEFAULT_TEST_SETTING_PREFIX
End Sub

Public Property Let Worksheet(ws As Worksheet)
    Set pWorksheet = ws
End Property

Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property


Public Property Let UseTestValues(use As Boolean)
    If use <> pUseTestValues Then
        Call ClearCache
    End If
    pUseTestValues = use
End Property

Public Property Get UseTestValues() As Boolean
    UseTestValues = pUseTestValues
End Property


Public Property Let TestSettingPrefix(prefix As String)
    pTestSettingPrefix = prefix
End Property

Public Property Get TestSettingPrefix() As String
    TestSettingPrefix = pTestSettingPrefix
End Property


Public Property Let CaseSensitive(sensitive As Boolean)
    pCaseSensitive = sensitive
End Property

Public Property Get CaseSensitive() As Boolean
    CaseSensitive = pCaseSensitive
End Property


Public Property Let CacheEnabled(enabled As Boolean)
    pCacheEnabled = enabled
    If Not enabled Then
        Call ClearCache
    End If
End Property

Public Property Get CacheEnabled() As Boolean
    CacheEnabled = pCacheEnabled
End Property


Public Function GetRangeByAddress(ByVal setting As String, ws As Worksheet) As Range
    Set GetRangeByAddress = ws.Range(GetSetting(setting).Value)
End Function

Public Function GetCellByPosition(ByVal setting As String, ws As Worksheet) As Range
    ' Setting must exist in the settings sheet as two cells, side by side, to specify row and column.
    Dim values As Range
    Set values = GetSetting(setting)
    
    On Error GoTo handler
    Set GetCellByPosition = ws.Cells(CLng(values.Cells(1).Value), CLng(values.Cells(2).Value))
    Exit Function
    
handler:
    Err.Description = "Invalid value for setting """ & setting & """. Must contain two numbers (row and column) on the same row."
    Err.Raise 1
End Function

Public Function GetWorksheetByName(ByVal setting As String, Optional wb As Workbook = Nothing)
    If wb Is Nothing Then: Set wb = ThisWorkbook
    Dim name As String
    name = GetSetting(setting)
    
    On Error GoTo handler
    Set GetWorksheetByName = wb.Sheets(name)
    Exit Function
    
handler:
    Err.Description = "Worksheet """ & name & """ not found in " & wb.name
    Err.Raise 1
End Function

Public Function GetCollection(ByVal setting As String) As Collection
    Dim rng As Range, cell As Range
    Set GetCollection = New Collection
    Set rng = GetSetting(setting)
    For Each cell In rng
        GetCollection.Add cell.Value
    Next cell
End Function

Public Function GetStringSet(ByVal setting As String) As Object 'Scripting.Dictionary
    Set GetStringSet = CreateObject("Scripting.Dictionary")
    Dim strings As Collection
    Set strings = GetStringList(setting)
    Dim s
    For Each s In strings
        If Not GetStringSet.exists(s) Then
            GetStringSet.Add key:=s, Item:=Nothing
        End If
    Next s
End Function

Public Function GetMapping(ByVal setting As String) As Object 'Scripting.Dictionary
    ' Setting range must be organized as a table of two columns
    ' Keys on the left, values on the right
    Set GetMapping = CreateObject("Scripting.Dictionary")
    Dim rng As Range
    Set rng = GetSetting(setting)
    
    If rng.Columns.Count < 2 Then
        Err.Description = "Setting """ & setting & """ must contain at least two columns for GetMapping to work."
        Err.Raise 1
    End If
    
    Dim row As Long
    For row = 1 To rng.Rows.Count
        Dim key, val
        key = rng.Cells(row, 1).Value
        val = rng.Cells(row, 2).Value
        If Not GetMapping.exists(key) Then
            GetMapping.Add key:=key, Item:=val
        End If
    Next row
End Function

Public Function GetSetting(ByVal setting As String) As Range
    Dim searchResult As Range
    Dim topLeft As Range
    Dim maxRow As Long, maxCol As Long
    Dim rightMostPosition As Long
    
    ' Check for a cache hit
    If CacheEnabled Then
        Set GetSetting = GetSettingFromCache(setting)
        If Not GetSetting Is Nothing Then
            Exit Function
        End If
    End If
    
    'find the setting name in the first column
    Set searchResult = FindSettingTitle(setting)
    If searchResult Is Nothing Then
        Err.Description = "Setting """ & setting & """ was not found in sheet """ & pWorksheet.name & """."
        Err.Raise 1
    End If
    
    'the top-left value of the setting is directly to the right of the name
    Set topLeft = searchResult.Offset(0, 1)
    'increment maxRow until we hit the either the next setting or an entirely blank row
    'keep track of the right-most used cell as well so we know how wide the table is
    maxRow = topLeft.row
    rightMostPosition = pWorksheet.Cells(maxRow, pWorksheet.Columns.Count).End(xlToLeft).Column
    maxCol = rightMostPosition
    
    Do Until pWorksheet.Cells(maxRow + 1, pWorksheet.Columns.Count).End(xlToLeft).Column = 1 Or pWorksheet.Cells(maxRow + 1, 1).Value <> vbNullString
        rightMostPosition = pWorksheet.Cells(maxRow + 1, pWorksheet.Columns.Count).End(xlToLeft).Column
        maxRow = maxRow + 1
        If rightMostPosition > maxCol Then
            maxCol = rightMostPosition
        End If
    Loop
    
    ' Compute the range of the setting
    Set GetSetting = Range(topLeft, pWorksheet.Cells(maxRow, maxCol))
    
    ' Add setting to cache
    If pCacheEnabled Then
        Dim cachekey As String
        cachekey = GetCacheKey(setting)
        pCache.Add key:=cachekey, Item:=GetSetting
    End If
    
End Function

Public Function HasSetting(setting As String) As Boolean
    If CacheEnabled Then
        If Not GetSettingFromCache(setting) Is Nothing Then
            HasSetting = True
            Exit Function
        End If
    End If
    HasSetting = Not (FindSettingTitle(setting) Is Nothing)
End Function

Private Function FindSettingTitle(ByVal setting As String) As Range
    ' Finds the position of the setting within the first column of the settings sheet

    ' If test mode is enabled, try to find the test value first
    Dim searchValue As String
    If pUseTestValues Then
        searchValue = GetTestSettingName(setting)
        Set FindSettingTitle = pWorksheet.Columns(1).Find(what:=searchValue, LookIn:=xlValues, lookat:=xlWhole, searchorder:=xlNext, MatchCase:=pCaseSensitive)
        If Not (FindSettingTitle Is Nothing) Then
            Exit Function
        End If
    End If
    
    Set FindSettingTitle = pWorksheet.Columns(1).Find(what:=setting, LookIn:=xlValues, lookat:=xlWhole, searchorder:=xlNext, MatchCase:=pCaseSensitive)
    
End Function

Private Function GetSettingFromCache(setting As String) As Range
    ' Check for the setting in the cache. Returns Nothing when the setting is not in cache.
    
    If pCache Is Nothing Then
        Set pCache = CreateObject("Scripting.Dictionary")
    End If
    
    ' Find the non-test value
    Dim key As String
    key = GetCacheKey(setting)
    If pCache.exists(key) Then
        Set GetSettingFromCache = pCache(key)
    End If
    
End Function

Private Function GetCacheKey(setting As String) As String
    GetCacheKey = pWorksheet.Parent.name & ":" & pWorksheet.name & "!" & setting
End Function

Public Function ClearCache()
    Set pCache = CreateObject("Scripting.Dictionary")
End Function

Private Function GetTestSettingName(setting As String) As String
    GetTestSettingName = pTestSettingPrefix & setting
End Function



