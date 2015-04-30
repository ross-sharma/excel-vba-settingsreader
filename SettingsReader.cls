Option Explicit

Public Sub Demo_GetSetting()

    ' Create a SettingsReader object
    Dim settings As New SettingsReader
    
    ' Set the settings sheet
    settings.Worksheet = ThisWorkbook.Sheets("Settings")
    
    
    ' Get some settings
    ' Note: GetSetting returns a Range
    settings.Worksheet.Activate
    Dim rng As Range
    
    Set rng = settings.GetSetting("server")
    rng.Select
    MsgBox "Setting ""server"" found at " & rng.Address
    
    Set rng = settings.GetSetting("port")
    rng.Select
    MsgBox "Setting ""port"" found at " & rng.Address
    
    Set rng = settings.GetSetting("users")
    rng.Select
    MsgBox "Setting ""users"" found at " & rng.Address
    

End Sub

Public Sub Demo_UseTestSettings()

    ' Create a SettingsReader object
    Dim settings As New SettingsReader
    
    ' Set the settings sheet
    settings.Worksheet = ThisWorkbook.Sheets("Settings")
    
    
    ' Enable test settings
    settings.UseTestValues = True
    ' The SettingsReader will now look for the prefix "_test_" when looking for settings
    ' You can change the prefix with the TestSettingPrefix property.
    
    ' Get some settings
    ' Note: GetSetting returns a Range
    settings.Worksheet.Activate
    Dim rng As Range
    
    Set rng = settings.GetSetting("server")
    rng.Select
    MsgBox "Setting ""server"" found at " & rng.Address & ". Notice that it used the test value."
    
    Set rng = settings.GetSetting("port")
    rng.Select
    MsgBox "Setting ""port"" foundat " & rng.Address & ". It didn't find a test value, so it returned the normal value."
    
    Set rng = settings.GetSetting("users")
    rng.Select
    MsgBox "Setting ""users"" found at " & rng.Address & ". It didn't find a test value, so it returned the normal value."
End Sub

Public Sub Demo_CaseSensitive()
    ' Use the CaseSensitive property to enable or disable case sensitivity
    ' CaseSensitive is False by default

    ' Create a SettingsReader object
    Dim settings As New SettingsReader
    
    ' Set the settings sheet
    settings.Worksheet = ThisWorkbook.Sheets("Settings")
    
    ' Get some settings
    ' Note: GetSetting returns a Range
    settings.Worksheet.Activate
    Dim rng As Range
    
    Set rng = settings.GetSetting("casesensitive")
    rng.Select
    MsgBox "Setting ""casesensitive"" found at " & rng.AddressLocal & ". CaseSensitive is FALSE."
    
    ' Enable case-sensitivity
    settings.CaseSensitive = True
    Set rng = settings.GetSetting("casesensitive")
    rng.Select
    MsgBox "Setting ""casesensitive"" found at " & rng.AddressLocal & ". CaseSensitive is now TRUE."
    
End Sub

Public Sub Demo_GetRangeByAddress()

    ' Create a SettingsReader object
    Dim settings As New SettingsReader
    
    ' Set the settings sheet
    settings.Worksheet = ThisWorkbook.Sheets("Settings")
    
    
    ' Note: GetRangeByAddress returns a Range
    Dim rng As Range
    Dim dataSheet As Worksheet: Set dataSheet = ThisWorkbook.Sheets("Data")
    
    Set rng = settings.GetRangeByAddress("some address", dataSheet)
    rng.Worksheet.Activate
    rng.Select
    MsgBox "GetRangeByAddress returned " & rng.Address

    settings.Worksheet.Activate
End Sub

Public Sub Demo_GetCellByPosition()

    ' Create a SettingsReader object
    Dim settings As New SettingsReader
    
    ' Set the settings sheet
    settings.Worksheet = ThisWorkbook.Sheets("Settings")
    
    ' Note: Demo_GetCellByPosition returns a single cell (a Range object)
    Dim rng As Range
    Dim dataSheet As Worksheet: Set dataSheet = ThisWorkbook.Sheets("Data")
    
    Set rng = settings.GetCellByPosition("some position", dataSheet)
    rng.Worksheet.Activate
    rng.Select
    MsgBox "GetCellByPosition returned " & rng.Address

    settings.Worksheet.Activate
End Sub

Public Sub Demo_GetWorksheetByName()

    ' Create a SettingsReader object
    Dim settings As New SettingsReader
    
    ' Set the settings sheet
    settings.Worksheet = ThisWorkbook.Sheets("Settings")
    
    
    ' Note: Demo_GetCellByPosition returns a single cell (a Range object)
    Dim rng As Range
    Dim dataSheet As Worksheet
    Set dataSheet = settings.GetWorksheetByName("some worksheet name")
    
    dataSheet.Activate
    MsgBox "GetWorksheetByName returned the worksheet named " & dataSheet.name & "."
    
    settings.Worksheet.Activate
End Sub

Public Sub Demo_GetCollection()
    ' Create a SettingsReader object
    Dim settings As New SettingsReader
    
    ' Set the settings sheet
    settings.Worksheet = ThisWorkbook.Sheets("Settings")
    
    ' Get the list of banned users as a Collection
    Dim strings As Collection
    Set strings = settings.GetCollection("banned users")
    
    Dim msg As String
    msg = "GetCollection(""Banned users"") returned a Collection containing " & strings.Count & " strings:" & vbNewLine
    Dim s
    For Each s In strings
        msg = msg & vbNewLine & s
    Next s
    MsgBox msg
End Sub

Public Sub Demo_GetMapping()
    ' Create a SettingsReader object
    Dim settings As New SettingsReader
    
    ' Set the settings sheet
    settings.Worksheet = ThisWorkbook.Sheets("Settings")
    
    ' Get the list of banned users as a Scripting.Dictionary
    Dim users As Object
    Set users = settings.GetMapping("users")
    
    Dim msg As String
    msg = "GetMapping(""users"") returned a Dictionary containing " & users.Count & " pairs:" & vbNewLine
    Dim key
    For Each key In users.keys
        msg = msg & vbNewLine & key & " --> " & users(key)
    Next key
    MsgBox msg
End Sub

