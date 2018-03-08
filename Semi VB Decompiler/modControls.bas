Attribute VB_Name = "modControls"
'*********************************************
'modControls
'Copyright VisualBasicZone.com 2004 - 2005
'*********************************************

'Control Sepeartor Constatns
Public Const vbFormNewChildControl = 511 'FF01
Public Const vbFormExistingChildControl = 767 'FF02
Public Const vbFormChildControl = 1023 'FF03
Public Const vbFormEnd = 1279 'FF04
Public Const vbFormMenu = 1535 'FF05

Public pMain As CPropertyItem
'Control Header
Public Type ControlHeader
   'Length As Integer
    'unknown As Integer
    Length As Long
    cId As Byte 'Used To link events
    cName As String
    un2 As Byte
    cType As Byte
End Type
Public Type ArrayTestType
    Length As Integer
    uni As Byte
    arrayflag As Byte
End Type
Public Type ControlArrayHeader
    Length As Integer
    un1 As Byte
   ' Length As Long
    arrayflag As Integer
    cId As Byte
    un2 As Byte
    cName As String
    un3 As Byte
    cType As Byte
End Type
Public Type ControlSize
    clientLeft As Integer
    un1 As Integer
    clientTop As Integer
    un2 As Integer
    clientWidth As Integer
    un3 As Integer
    clientHeight As Integer
    un4 As Integer
End Type
'Used in cType
Public Enum ControlType
    vbPictureBox = 0
    vbLabel = 1
    vbTextBox = 2
    vbFrame = 3
    vbCommandbutton = 4
    vbCheckbox = 5
    vbOptionbutton = 6
    vbComboBox = 7
    vbListbox = 8
    vbHscroll = 9
    vbVscroll = 10
    vbTimer = 11
    vbform = 13
    vbDriveListbox = 16
    vbDirectoryListbox = 17
    vbFileListbox = 18
    vbmenu = 19
    vbMDIForm = 20
    vbShape = 22
    vbLine = 23
    vbImage = 24
    vbData = 37
    vbOLE = 38
    vbUserControl = 40
    vbPropertyPage = 41
    vbUserDocument = 42
End Enum



'External Controls
Private Type OcxListType
    strGuid As String
    strocxName As String
    strLibname As String
    strName As String
End Type

Global gOcxList() As OcxListType


Public Type FontType
    un1 As Byte
    un2 As Byte
    un3 As Byte
    action As Byte
    Weight As Integer
    Size As Long
    FontLen As Byte
End Type


Public Type tControlEventLink

    Const1 As Integer        ' 0x00
    CompileType As Byte      ' 0x02 compileType According to Sarge[more info?]
    aEvent As Long           ' 0x03
                             ' 0x07 &lt;-- Structure Size
End Type


Public Type tControlEventPointer
    Const1 As Byte          ' 0x00
    Flag1 As Long           ' 0x01
    Const2 As Integer       ' 0x05
    EventLink As tControlEventLink ' 0x07
                            ' 0x0E &lt;-- Structure Size
End Type

Public Type LineSizeType
    X1 As Long
    op1 As Byte
    Y1 As Long
    op2 As Byte
    X2 As Long
    op3 As Byte
    Y2 As Long
End Type


Dim TempLength As Double
Dim TempText As String
Dim pointer As Double

Private Type Data_Format
    DataType As Integer
    DataFormat As String
    DataHaveTFNull As Integer
    DataTrueValue As String
    DataFalseValue As String
    DataNullValue As String
    DataFDOW As Integer 'First day of week
    DataFWOY As Integer 'First week of year
    DataLCID As Integer
    DataSubType As Integer
End Type

Private DataFormat As Data_Format
Global IdentNextMenu As Boolean
Public gComboBoxStyle As Boolean
Public Sub GetDataFormat(ByVal F As Integer, Optional bEditor As Boolean = False)
On Error GoTo errHandle:
    Dim TempDouble2 As Double
    Dim TempLength1 As Double
    Dim TempLength2 As Double
    Dim TempLength3 As Double
    Dim TempString As String
    Dim LoopCounter As Integer
    Dim TempChar As String
        
    'Skip 4 byte text "P8&k"
    TempDouble# = GetByteByFile(F)
    TempDouble# = GetByteByFile(F)
    TempDouble# = GetByteByFile(F)
    TempDouble# = GetByteByFile(F)
        
    'Skip 2 words
    TempDouble# = GetWordByFile(F)   '2
    TempDouble# = GetWordByFile(F)  '6
    
    'Get type
    TempDouble# = GetWordByFile(F)
    
    'Save it
    DataFormat.DataType = TempDouble#
    
    'Is it a BOOLEAN data type?
    If DataFormat.DataType = DATA_BOOLEAN_TYPE Then
    
        'Skip a word
        TempDouble# = GetWordByFile(F) '0
        
        'Get day of week
        TempDouble# = GetWordByFile(F) '1  5  FirstDayOfWeek
        
        'Save it
        DataFormat.DataFDOW = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F) '0
        
        'Get week of year
        TempDouble# = GetWord() '2  4  FirstWeekOfYear
        
        'Save it
        DataFormat.DataFWOY = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F) '0
        
        'Get format length
        TempDouble# = GetWordByFile(F)
        
        'Save it
        TempLength = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F)
        
        'Get TrueFalseNull
        TempDouble# = GetWordByFile(F)
        
        'Save it
        DataFormat.DataHaveTFNull = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F)
        
        'Get first text length
        TempLength1 = GetWordByFile(F)
        
        'Skip a word
        TempDouble# = GetWordByFile(F)
        
        'Get second text length
        TempLength2 = GetWordByFile(F)
        
        'Skip a word
        TempDouble# = GetWordByFile(F)
        
        'Get third text length
        TempLength3 = GetWordByFile(F)
        
        'Skip a word
        TempDouble# = GetWordByFile(F)
        
        'Get format text
        DataFormat.DataFormat = GetUnicodeStringWLen(TempLength, F)
       
        'Skip 8 words()
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        
        'Get TRUE text
        DataFormat.DataTrueValue = GetUnicodeStringWLen(TempLength1, F)
          
        'Skip 8 words()
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
          
        'Get FALSE text
        DataFormat.DataFalseValue = GetUnicodeStringWLen(TempLength2, F)
         
        'Skip 8 words()
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
          
        'Get NULL text
        DataFormat.DataNullValue = GetUnicodeStringWLen(TempLength3, F)
         
        'Get LCID
        TempDouble# = GetWordByFile(F)
        
        'Save it
        DataFormat.DataLCID = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F)
        
        'Get subtype
        TempDouble# = GetWordByFile(F)
        
        'Save it
        DataFormat.DataSubType = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F)
        
    'Is it a CHECKBOX data type?
    ElseIf DataFormat.DataType = DATA_CHECKBOX_TYPE Then
    
         'Skip a word
        TempDouble# = GetWordByFile(F) '0
        
        'Get day of week
        TempDouble# = GetWordByFile(F) '1  5  FirstDayOfWeek
        
        'Save it
        DataFormat.DataFDOW = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F) '0
        
        'Get week of year
        TempDouble# = GetWordByFile(F) '2  4  FirstWeekOfYear
        
        'Save it
        DataFormat.DataFWOY = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F) '0
        
        'Skip format length of 0x00
        TempDouble# = GetWordByFile(F)
                
        'Skip 3 words
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        TempDouble# = GetWordByFile(F)
        
       'Get text
        TempText$ = GetUnicodeStringWLen(6, F)
    
        'Get LCID
        TempDouble# = GetWordByFile(F)
        
        'Save it
        DataFormat.DataLCID = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F)
        
        'Get subtype
        TempDouble# = GetWordByFile(F)
        
        'Save it
        DataFormat.DataSubType = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F)
    
        
    'Must be other type
    Else
        
        'Skip a word
        TempDouble# = GetWordByFile(F) '0
        
        'Get day of week
        TempDouble# = GetWordByFile(F) '1  5  FirstDayOfWeek
        
        'Save it
        DataFormat.DataFDOW = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F) '0
        
        'Get week of year
        TempDouble# = GetWordByFile(F) '2  4  FirstWeekOfYear
        
        'Save it
        DataFormat.DataFWOY = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F) '0
        
        'Get format length
        TempDouble# = GetWordByFile(F)
        
        'Save it
        TempLength = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F)
        
        'Get TrueFalseNull
        TempDouble# = GetWordByFile(F)
        
        'Save it
        DataFormat.DataHaveTFNull = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F)
    
        'Get text
        TempText$ = GetUnicodeStringWLen(6, F)
        
        'Get format text, if any
        If TempLength > 0 Then
        
            'Get the string
            TempString$ = GetUnicodeStringWLen(TempLength, F)
        
            'Clear the existing data
            DataFormat.DataFormat = ""
            
            'Loop throught text
            For LoopCounter% = 1 To TempLength
            
                'Get each character
                TempChar$ = Mid$(TempString$, LoopCounter%, 1)
                
                'Check for imbedded quote
                If TempChar$ = Chr$(34) Then
                    'Add second quote
                    DataFormat.DataFormat$ = DataFormat.DataFormat$ & Chr$(34) & TempChar$
                Else
                    'Keep original
                    DataFormat.DataFormat$ = DataFormat.DataFormat$ & TempChar$
                End If
                
            Next
        
        End If
        
        'Get LCID
        TempDouble# = GetWordByFile(F)
        
        'Save it
        DataFormat.DataLCID = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F)
        
        'Get subtype
        TempDouble# = GetWordByFile(F)
        
        'Save it
        DataFormat.DataSubType = TempDouble#
        
        'Skip a word
        TempDouble# = GetWordByFile(F)
    
    End If
    If bEditor = False Then
        
        gIdentSpaces = gIdentSpaces + 1
        Call AddText("BeginProperty DataFormat")
        gIdentSpaces = gIdentSpaces + 1
    
                    
        Call AddText("Type = " & DataFormat.DataType)
        Call AddText("Format =" & cQuote & DataFormat.DataFormat & cQuote)
        Call AddText("HaveTrueFalseNull = " & DataFormat.DataHaveTFNull)
        
        Call AddText("TrueValue = " & cQuote & DataFormat.DataTrueValue & cQuote)
        Call AddText("FalseValue = " & cQuote & DataFormat.DataFalseValue & cQuote)
        Call AddText("NullValue = " & cQuote & DataFormat.DataNullValue & cQuote)
        
        Call AddText("FirstDayOfWeek = " & DataFormat.DataFDOW)
        Call AddText("FirstWeekOfYear = " & DataFormat.DataFWOY)
        Call AddText("LCID = " & DataFormat.DataLCID)
        Call AddText("SubFormatType = " & DataFormat.DataSubType)
        
        gIdentSpaces = gIdentSpaces - 1
        Call AddText("EndProperty")
        gIdentSpaces = gIdentSpaces - 1
    End If
Exit Sub
errHandle:
    Call modGlobals.AddToErrorLog("GetDataFormat: " & err.Description)
    
End Sub
Public Function GetUnicodeStringWLen(ByVal TextLen As Double, ByVal F As Integer) As String

    Dim count As Double
    Dim bData As Byte
    'Clear string
    GetUnicodeStringWLen$ = ""
 
    'Get a Unicode string char by char
    For count = 1 To TextLen
        Get #F, , bData
        'Get a char
        GetUnicodeStringWLen$ = GetUnicodeStringWLen$ & Chr$(bData)
        
        'Skip the 0 byte
        Seek #F, Seek(F) + 1
        
    Next
    
End Function
'##################################
'Begin Subs for Processing Special opcodes and properties for common controls
'##################################
Public Function ProccessForm(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
        Dim tByte As Byte
        Dim iData As Integer
If gVB4App = True Then
    If bCode = 5 Then
        '53 is the size opcode for form's
        Dim objectSize As ControlSize
        Get F, , objectSize

        Call AddText("ClientLeft = " & objectSize.clientLeft)
        Call AddText("ClientTop = " & objectSize.clientTop)
        Call AddText("ClientWidth = " & objectSize.clientWidth)
        Call AddText("ClientHeight = " & objectSize.clientHeight)

        ProccessForm = 16
        Exit Function
    End If

    If bCode = 53 Then
        Call AddText("Left = " & GetLong(F))
        ProccessForm = 4
        Exit Function
    End If
    If bCode = 54 Then
        Call AddText("Top = " & GetLong(F))
        ProccessForm = 4
        Exit Function
    End If
    If bCode = 55 Then
        Call AddText("ScaleWidth = " & GetLong(F))
        ProccessForm = 4
        Exit Function
    End If
    If bCode = 56 Then
        Call AddText("ScaleHeight = " & GetLong(F))
        ProccessForm = 4
        Exit Function
    End If
End If
    If bCode = 10 Then
        Call AddText("WindowState = " & GetByte2(F))
        ProccessForm = 1
        Exit Function
    End If
    'DrawStyle
    If bCode = 27 Then
        Call AddText("DrawStyle = " & GetByte2(F))
        ProccessForm = 1
        Exit Function
    End If
    'FillStyle
    If bCode = 29 Then
        Call AddText("FillStyle = " & GetByte2(F))
        ProccessForm = 1
        Exit Function
    End If
    'LinkMode
    If bCode = 37 Then
        Call AddText("LinkMode = " & GetByte2(F))
        ProccessForm = 1
        Exit Function
    End If
    
    If bCode = 53 And gVB4App = False Then
        '53 is the size opcode for form's
        Dim objectSize2 As ControlSize
            Get F, , objectSize2
            If bEditor = False Then
                Call AddText("ClientLeft = " & objectSize2.clientLeft)
                Call AddText("ClientTop = " & objectSize2.clientTop)
                Call AddText("ClientWidth = " & objectSize2.clientWidth)
                Call AddText("ClientHeight = " & objectSize2.clientHeight)
            Else
            
                Call modGlobals.AddPropertyToTheList("ClientLeft", objectSize2.clientLeft, "Long", Loc(F) - 16, "Returns/sets the distance between the internal left edge of an object and the left edge of its container.")
                Call modGlobals.AddPropertyToTheList("ClientTop", objectSize2.clientTop, "Long", Loc(F) - 12, "Returns/sets the distance between the internal top edge of an object and the top edge of its container.")
                Call modGlobals.AddPropertyToTheList("ClientWidth", objectSize2.clientWidth, "Long", Loc(F) - 8, "Returns/sets the width of an object.")
                Call modGlobals.AddPropertyToTheList("ClientHeight", objectSize2.clientHeight, "Long", Loc(F) - 4, "Returns/sets the height of an object.")
            End If
        
        ProccessForm = 16
        Exit Function
    End If
    
    'Form Flags
    If bCode = 25 Then
        Dim byteScaleMode As Byte
        'Get 1st flag byte
         tByte = GetByte2(F)
        'Set scale mode
         byteScaleMode = tByte
         Call AddText("ScaleMode = " & byteScaleMode)
        'Get 2nd flag byte
         tByte = GetByte2(F)
                
        'Does ScaleMode = User?
        If byteScaleMode = &H0 Then
            'Skip next 16 bytes
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
            tByte = GetByte2(F)
                                        
        End If
                            
        'Get 3rd flag byte
        tByte = GetByte2(F)
        Dim bRedraw As Boolean
        Dim bFontTrans As Boolean
        'Derive auto redraw
        bRedraw = ((tByte And &H20) <> 0)
                
        'Derive font transparent
        bFontTrans = ((tByte And &H2) <> 0)
        If bRedraw = True Then
            Call AddText("AutoRedraw = -1")
        Else
            Call AddText("AutoRedraw = 0")
        End If
        If bFontTrans = True Then
            Call AddText("FontTransparent = -1")
        Else
            Call AddText("FontTransparent = 0")
        End If

        'Get 4th flag byte
        tByte = GetByte2(F)
        ProccessForm = 1
        Exit Function
    End If
    
    If bCode = 64 Then
        'Font
        Call frmMain.GetFontProperty(F)
         ProccessForm = 11
        Exit Function
    End If

    
    If bCode = 61 Then

        tByte = GetByte2(F)
        If tByte = 255 Then
            iData = -1
        Else
            iData = 0
        End If
        Call AddText("LockControls = " & iData)
        ProccessForm = 1
        Exit Function
    End If
    If bCode = 62 Then

        tByte = GetByte2(F)
        If tByte = 255 Then
            iData = -1
        Else
            iData = 0
        End If
        Call AddText("NegotiateMenus = " & iData)
        ProccessForm = 1
        Exit Function
    End If
    If bCode = 65 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessForm = 1
        Exit Function
    End If
    If bCode = 11 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessForm = 1
        Exit Function
    End If
    If bCode = 31 Then
        Call AddText("DrawMode = " & GetByte2(F))
        ProccessForm = 1
        Exit Function
    End If
    If bCode = 34 Then
        Call AddText("BorderStyle = " & GetByte2(F))
        ProccessForm = 1
        Exit Function
    End If
    If bCode = 70 Then
        Call AddText("StartUpPosition = " & GetByte2(F))
        ProccessForm = 1
        Exit Function
    End If
    If bCode = 71 Then
        Call AddText("OLEDropMode =" & GetByte2(F))
        ProccessForm = 1
        Exit Function
    End If
    If bCode = 73 Then
        Call AddText("PaletteMode = " & GetByte2(F))
        ProccessForm = 1
        Exit Function
    End If
    If bCode = 98 Then
        GetByte2 (F)
        ProccessForm = 1
        Exit Function
    End If
    If bCode = 99 Then
        GetByte2 (F)
        ProccessForm = 1
        Exit Function
    End If
    If bCode = 0 Then 'Null Byte
        GetByte2 (F)
        ProccessForm = 1
        Exit Function
    End If
    ProccessForm = -1
End Function

Public Function ProccessCheckBox(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer

    If bCode = 5 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessCheckBox = 8
        Exit Function
    End If
    If bCode = 19 Then
        Call AddText("Value = " & GetByte2(F))
        ProccessCheckBox = 1
        Exit Function
    End If
    If bCode = 25 Then
        Call AddText("Alignment = " & GetByte2(F))
        ProccessCheckBox = 1
        Exit Function
    End If
    'DataSource
    If bCode = 28 Then
        Dim dataLength As Integer
        dataLength = GetInteger(F)
        Call AddText("DataSource = " & cQuote & GetAllString(F) & cQuote)
        ProccessCheckBox = 2
        Exit Function
    End If
    If bCode = 32 Then
        Call frmMain.GetFontProperty(F)
        ProccessCheckBox = 11
        Exit Function
    End If
    If bCode = 34 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessCheckBox = 1
        Exit Function
    End If
    If bCode = 41 Then
        Call AddText("Style = " & GetByte2(F))
        ProccessCheckBox = 1
        Exit Function
    End If
    If bCode = 11 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessCheckBox = 1
        Exit Function
    End If
    If bCode = 21 Then
        Call AddText("DragMode =" & GetByte2(F))
        ProccessCheckBox = 1
        Exit Function
    End If
    If bCode = 42 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessCheckBox = 1
        Exit Function
    End If
    If bCode = 47 Then 'DataFormat
        Call modControls.GetDataFormat(F, bEditor)
        ProccessCheckBox = 1
        Exit Function
    End If
    ProccessCheckBox = -1
End Function

Public Function ProccessComboBox(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 1 Then
        Dim hB As Byte
        hB = GetByte2(F)
        Call AddText("Style = " & hB)
        If hB = 2 Then
            gComboBoxStyle = True
        Else
            gComboBoxStyle = False
        End If
         ProccessComboBox = 1
        Exit Function
    End If
    If bCode = 5 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessComboBox = 8
        Exit Function
    End If
    If bCode = 0 Then
        Call GetByte2(F)
        ProccessComboBox = 1
        Exit Function
    End If
    If bCode = 11 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessComboBox = 1
        Exit Function
    End If
    If bCode = 22 Then
        ProccessComboBox = frmMain.GetListType(F)
        Exit Function
    End If
    'Text
    If bCode = 12 And gComboBoxStyle = True Then
        ProccessComboBox = 0
        Exit Function
    End If
    'DragMode
    If bCode = 28 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessComboBox = 1
        Exit Function
        
    End If
    If bCode = 43 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessComboBox = 1
        Exit Function
    End If
    If bCode = 33 Then
        Dim TempDouble As Double
        Dim TempInt As Integer
        TempDouble = GetWordByFile(CInt(F))
        If TempDouble <> 0 Then
            TempInt = GetByte2(F)
            TempInt = GetByte2(F)
                    Do
                    
                        'Get length of item
                        TempInt% = GetWordByFile(CInt(F))
                                  
                        'Loop through "TempInt" chars
                        Do
                        
                            'Get a char
                            'ComboData.ItemData = ComboData.ItemData & Chr$(GetByte())
                            GetByte2 (F)
                            'Reset text length
                            TempInt% = TempInt% - 1
                            
                        Loop While TempInt% <> 0
                        
                        'Add in a CR
                        'ComboData.ItemData = ComboData.ItemData & Chr$(13)
                        
                        'Reset item count
                        TempDouble# = TempDouble# - 1
                            
                    Loop While TempDouble# > 0
                    
        End If
        ProccessComboBox = 1
        Exit Function
    End If
    If bCode = 38 Then
        Call frmMain.GetFontProperty(F)
        ProccessComboBox = 11
        Exit Function
    End If
    'DataSource
    If bCode = 39 Then
        Dim dataLength As Integer
        dataLength = GetInteger(F)
        Call AddText("DataSource = " & cQuote & GetAllString(F) & cQuote)
        ProccessComboBox = 2
        Exit Function
    End If
    'OLEDragMode
    If bCode = 48 Then
        Call AddText("OLEDragMode = " & GetByte2(F))
        ProccessComboBox = 1
        Exit Function
    End If
    'OLEDropMode
    If bCode = 49 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessComboBox = 1
        Exit Function
    End If
    If bCode = 54 Then 'DataFormat
        Call modControls.GetDataFormat(F, bEditor)
        ProccessComboBox = 1
        Exit Function
    End If

    ProccessComboBox = -1
End Function

Public Function ProccessCommandButton(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 4 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessCommandButton = 8
        Exit Function
    End If
    If bCode = 29 Then
        Call frmMain.GetFontProperty(F)
        ProccessCommandButton = 11
        Exit Function
    End If
    If bCode = 31 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessCommandButton = 1
        Exit Function
    End If
    If bCode = 41 Then
        Call AddText("Style = " & GetByte2(F))
        ProccessCommandButton = 1
        Exit Function
    End If
    If bCode = 10 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessCommandButton = 1
        Exit Function
    End If
    If bCode = 22 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessCommandButton = 1
        Exit Function
    End If
    If bCode = 38 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessCommandButton = 1
        Exit Function
    End If
    ProccessCommandButton = -1
End Function
Public Function ProccessDataControl(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    'Control Postion opcode
    If bCode = 2 Then
        Call frmMain.GetControlSize(F)
        ProccessDataControl = 8
        Exit Function
    End If
    If bCode = 8 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessDataControl = 1
        Exit Function
    End If
    If bCode = 24 Then
        Dim iData As Integer
        iData = GetInteger(F)
        If iData = 0 Then
            Call AddText("RecordSource = " & cQuote & cQuote)
        Else
            Call AddText("RecordSource = " & cQuote & GetAllString(F) & cQuote)
        End If
        ProccessDataControl = 2
        Exit Function
    End If
    'DragMode
    If bCode = 30 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessDataControl = 1
        Exit Function
    End If
    If bCode = 33 Then
        Call frmMain.GetFontProperty(F)
        ProccessDataControl = 1
        Exit Function
    End If
    'BOFAction
    If bCode = 35 Then
        Call AddText("BOFAction = " & GetByte2(F))
        ProccessDataControl = 1
        Exit Function
    End If
    'EOFAction
    If bCode = 36 Then
        Call AddText("EOFAction = " & GetByte2(F))
        ProccessDataControl = 1
        Exit Function
    End If
    If bCode = 37 Then
        Call AddText("RecordsetType = " & GetByte2(F))
        ProccessDataControl = 1
        Exit Function
    End If
    'Appearance
    If bCode = 39 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessDataControl = 1
        Exit Function
    End If
    'Negotiate
    If bCode = 40 Then
        GetByte2 (F)
        Call AddText("Negotiate = -1")
        ProccessDataControl = 1
        Exit Function
    End If
    If bCode = 42 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessDataControl = 1
        Exit Function
    End If
    If bCode = 44 Then
        Call AddText("DefaultType = " & GetByte2(F))
        ProccessDataControl = 1
        Exit Function
    End If
    If bCode = 45 Then
        Call AddText("DefaultCursorType = " & GetByte2(F))
        ProccessDataControl = 1
        Exit Function
    End If


    ProccessDataControl = -1
End Function
Public Function ProccessDesigner(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer

    ProccessDesigner = -1
End Function

Public Function ProccessDirListBox(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    'Control Postion opcode
    If bCode = 4 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessDirListBox = 8
        Exit Function
    End If
    If bCode = 10 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessDirListBox = 1
        Exit Function
    End If
    If bCode = 23 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessDirListBox = 1
        Exit Function
    End If
    If bCode = 33 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessDirListBox = 1
        Exit Function
    End If
    If bCode = 36 Then
        Call AddText("OLEDragMode = " & GetByte2(F))
        ProccessDirListBox = 1
        Exit Function
    End If
    If bCode = 37 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessDirListBox = 1
        Exit Function
    End If
    If bCode = 31 Then
        Call frmMain.GetFontProperty(F)
        ProccessDirListBox = 11
        Exit Function
    End If
    ProccessDirListBox = -1
End Function
Public Function ProccessDriveListBox(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 4 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessDriveListBox = 8
        Exit Function
    End If
    If bCode = 23 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessDriveListBox = 1
        Exit Function
    End If
    If bCode = 30 Then
        Call frmMain.GetFontProperty(F)
        ProccessDriveListBox = 11
        Exit Function
    End If
    If bCode = 10 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessDriveListBox = 1
        Exit Function
    End If
    If bCode = 32 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessDriveListBox = 1
        Exit Function
    End If
    If bCode = 35 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessDriveListBox = 1
        Exit Function
    End If
    ProccessDriveListBox = -1
End Function
Public Function ProccessFileListBox(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 4 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessFileListBox = 8
        Exit Function
    End If
    If bCode = 41 Then
        Call frmMain.GetFontProperty(F)
        ProccessFileListBox = 11
        Exit Function
    End If
    If bCode = 10 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessFileListBox = 1
        Exit Function
    End If
    If bCode = 30 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessFileListBox = 1
        Exit Function
    End If
    If bCode = 36 Then
        Call AddText("MultiSelect = " & GetByte2(F))
        ProccessFileListBox = 1
        Exit Function
    End If
    If bCode = 43 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessFileListBox = 1
        Exit Function
    End If
    If bCode = 46 Then
        Call AddText("OLEDragMode = " & GetByte2(F))
        ProccessFileListBox = 1
        Exit Function
    End If
    If bCode = 47 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessFileListBox = 1
        Exit Function
    End If
    ProccessFileListBox = -1
End Function
Public Function ProccessFrame(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    
    If bCode = 5 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessFrame = 8
        Exit Function
    End If
    If bCode = 4 Then
        Call AddText("ForeColor = " & GetLong(F))
        ProccessFrame = 4
        Exit Function
    End If
    If bCode = 27 Then
        Call frmMain.GetFontProperty(F)
        ProccessFrame = 11
        Exit Function
    End If
    If bCode = 34 Then
        Call AddText("BorderStyle = " & GetByte2(F))
        ProccessFrame = 1
        Exit Function
    End If
    If bCode = 18 Then
        Call AddText("TabIndex=" & GetInteger(F))

        ProccessFrame = 2
        Exit Function
    End If
    If bCode = 29 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessFrame = 1
        Exit Function
    End If
    If bCode = 11 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessFrame = 1
        Exit Function
    End If
    If bCode = 20 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessFrame = 1
        Exit Function
    End If
    If bCode = 33 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessFrame = 1
        Exit Function
    End If
    ProccessFrame = -1
End Function
Public Function ProccessHscroll(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 2 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessHscroll = 8
        Exit Function
    End If
    If bCode = 16 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessHscroll = 1
        Exit Function
    End If
    If bCode = 8 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessHscroll = 1
        Exit Function
    End If
    ProccessHscroll = -1
End Function

Public Function ProccessImage(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 3 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessImage = 8
        Exit Function
    End If
    If bCode = 15 Then
        Call AddText("BorderStyle = " & GetByte2(F))
        ProccessImage = 1
        Exit Function
    End If
    'DataSource
    If bCode = 16 Then
        Dim dataLength As Integer
        dataLength = GetInteger(F)
        Call AddText("DataSource = " & cQuote & GetAllString(F) & cQuote)
        ProccessImage = 2
        Exit Function
    End If
    If bCode = 21 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessImage = 1
        Exit Function
    End If
    If bCode = 12 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessImage = 1
        Exit Function
    End If
    If bCode = 24 Then
        Call AddText("OLEDragMode = " & GetByte2(F))
        ProccessImage = 1
        Exit Function
    End If
    If bCode = 25 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessImage = 1
        Exit Function
    End If
    If bCode = 9 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessImage = 1
        Exit Function
    End If
    If bCode = 27 Then 'DataFormat
        Call modControls.GetDataFormat(F, bEditor)
        ProccessImage = 1
        Exit Function
    End If
    ProccessImage = -1
End Function
Public Function ProccessLabel(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 5 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessLabel = 8
        Exit Function
    End If
    If bCode = 11 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessLabel = 1
        Exit Function
    End If
    If bCode = 19 Then
        Call AddText("BorderStyle = " & GetByte2(F))
        ProccessLabel = 1
        Exit Function
    End If
    If bCode = 20 Then
        Call AddText("Alignment = " & GetByte2(F))
        ProccessLabel = 1
        Exit Function
    End If
    If bCode = 31 Then
        Call AddText("BackStyle = " & GetByte2(F))
        ProccessLabel = 1
        Exit Function
    End If
    'DataSource
    If bCode = 32 Then
        Dim dataLength As Integer
        dataLength = GetInteger(F)
        Call AddText("DataSource = " & cQuote & GetAllString(F) & cQuote)
        ProccessLabel = 2
        Exit Function
    End If
    If bCode = 37 Then
        'Font
        Call frmMain.GetFontProperty(F)
        ProccessLabel = 11
        Exit Function
    End If
    If bCode = 39 Then ' Appearance
        Call AddText("Appearance = " & GetByte2(F))
        ProccessLabel = 1
        Exit Function
    End If
    If bCode = 26 Then
        Call AddText("DragMode =" & GetByte2(F))
        ProccessLabel = 1
        Exit Function
    End If
    If bCode = 43 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessLabel = 1
        Exit Function
    End If
    If bCode = 45 Then 'DataFormat
        Call modControls.GetDataFormat(F, bEditor)
        ProccessLabel = 1
        Exit Function
    End If
    ProccessLabel = -1
End Function
Public Function ProccessLine(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 3 Then
        Dim LineSize As LineSizeType
        Get F, , LineSize
        Call AddText("X1 = " & LineSize.X1)
        Call AddText("Y1 = " & LineSize.Y1)
        Call AddText("X2 = " & LineSize.X2)
        Call AddText("Y2 = " & LineSize.Y2)
        Exit Function
    End If
    If bCode = 11 Then
        Call AddText("BorderStyle = " & GetByte2(F))
        ProccessLine = 1
        Exit Function
    End If
    If bCode = 10 Then
        Call AddText("DrawMode = " & GetByte2(F))
        ProccessLine = 1
        Exit Function
    End If
    ProccessLine = -1
End Function
Public Function ProccessListBox(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 4 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessListBox = 8
        Exit Function
    End If

   ''If bCode = 20 Then
    '    Dim iNumOfItems As Integer
    '    iNumOfItems = GetInteger(f)
    '    junk = GetInteger(f)
    '    Dim i As Integer
    '    Dim iLength As Integer
    '    Dim strData As String
    '    For i = 1 To iNumOfItems
    '        iLength = GetInteger(f)
    '        'MsgBox Loc(F)
    '        strData = gVBFile.GetString(Loc(f), iLength, False)
    '        'Seek #F, Loc(F) + strLength + 2
    '        'MsgBox strData
    '    Next i
    '    ProccessListBox = 0
    '    Exit Function
    'End If
    If bCode = 20 Then
        ProccessListBox = frmMain.GetListType(F)
        Exit Function
    End If
    'DragMode
    If bCode = 24 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessListBox = 1
        Exit Function
    End If
    
    If bCode = 44 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessListBox = 1
        Exit Function
    End If
    If bCode = 10 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessListBox = 1
        Exit Function
    End If
    If bCode = 29 Then
        Call AddText("MultiSelect = " & GetByte2(F))
        ProccessListBox = 1
        Exit Function
    End If
    
    If bCode = 39 Then
        Call frmMain.GetFontProperty(F)
        ProccessListBox = 11
        Exit Function
    End If
    'DataSource
    If bCode = 40 Then
        Dim dataLength As Integer
        dataLength = GetInteger(F)
        Call AddText("DataSource = " & cQuote & GetAllString(F) & cQuote)
        ProccessListBox = 2
        Exit Function
    End If
    If bCode = 49 Then
        Call AddText("OLEDragMode = " & GetByte2(F))
        ProccessListBox = 1
        Exit Function
    End If
    If bCode = 50 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessListBox = 1
        Exit Function
    End If
    If bCode = 51 Then
        Call AddText("Style = " & GetByte2(F))
        ProccessListBox = 1
        Exit Function
    End If
    If bCode = 54 Then 'DataFormat
        Call modControls.GetDataFormat(F, bEditor)
        ProccessListBox = 1
        Exit Function
    End If
    ProccessListBox = -1
End Function
Public Function ProccessMenu(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 7 Then 'Ident
        GetByte2 (F)
        IdentNextMenu = True
        ProccessMenu = 1
        Exit Function
    End If
    
    If bCode = 6 Then 'Seperator
        GetByte2 (F)
        ProccessMenu = 1
        Exit Function
    End If
    If bCode = 8 Then
        Call AddText("Shortcut = " & GetInteger(F))
        ProccessMenu = 0
        Exit Function
    End If
        ProccessMenu = -1
End Function
Public Function ProccessOLE(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 3 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessOLE = 8
        Exit Function
    End If
    'DragMode
    If bCode = 9 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessOLE = 1
        Exit Function
    End If
    'BorderStyle
    If bCode = 12 Then
        Call AddText("BorderStyle = " & GetByte2(F))
        ProccessOLE = 1
        Exit Function
    End If
    'MousePointer
    If bCode = 16 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessOLE = 1
        Exit Function
    End If
    'OLEType
    If bCode = 18 Then
        Call AddText("OLEType = " & GetByte2(F))
        ProccessOLE = 1
        Exit Function
    End If
    'UpdateOptions
    If bCode = 21 Then
        Call AddText("UpdateOptions = " & GetByte2(F))
        ProccessOLE = 1
        Exit Function
    End If
    'SizeMode
    If bCode = 23 Then
        Call AddText("SizeMode = " & GetByte2(F))
        ProccessOLE = 1
        Exit Function
    End If
    'AutoActivate
    If bCode = 24 Then
        Call AddText("AutoActivate = " & GetByte2(F))
        ProccessOLE = 1
        Exit Function
    End If
    'OLETypeAllowed
    If bCode = 36 Then
        Call AddText("OLETypeAllowed = " & GetByte2(F))
        ProccessOLE = 1
        Exit Function
    End If
    ProccessOLE = -1
End Function
Public Function ProccessOption(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 5 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessOption = 8
        Exit Function
    End If
    If bCode = 29 Then
        Call frmMain.GetFontProperty(F)
        ProccessOption = 11
        Exit Function
    End If
    If bCode = 11 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessOption = 1
        Exit Function
    End If
    If bCode = 21 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessOption = 1
        Exit Function
    End If
    If bCode = 25 Then
        Call AddText("Alignment = " & GetByte2(F))
        ProccessOption = 1
        Exit Function
    End If
    If bCode = 31 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessOption = 1
        Exit Function
    End If
    If bCode = 38 Then
        Call AddText("Style = " & GetByte2(F))
        ProccessOption = 1
        Exit Function
    End If
    If bCode = 39 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessOption = 1
        Exit Function
    End If
    ProccessOption = -1
End Function
Public Function ProccessPictureBox(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 5 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessPictureBox = 8
        Exit Function
    End If
    If bCode = 11 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessPictureBox = 1
        Exit Function
    End If
    If bCode = 36 Then
        Call AddText("BorderStyle = " & GetByte2(F))
        ProccessPictureBox = 1
        Exit Function
    End If
    If bCode = 57 Then
        Call frmMain.GetFontProperty(F)
        ProccessPictureBox = 11
        Exit Function
    End If
    If bCode = 59 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessPictureBox = 1
        Exit Function
    End If
    If bCode = 30 Then
        Call AddText("FillStyle = " & GetByte2(F))
        ProccessPictureBox = 1
        Exit Function
    End If
    'Flags
    If bCode = 26 Then
        Dim bScaleMode As Byte
        bScaleMode = GetByte2(F)
        Call AddText("ScaleMode = " & bScaleMode)
        GetByte2 (F)
        If bScaleMode = 0 Then
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
            GetByte2 (F)
        End If
        Dim bFlag As Byte
        Dim bRedraw As Boolean
        Dim bFontTrans As Boolean
        bFlag = GetByte2(F)
        bRedraw = ((bFlag And &H20) <> 0)
        If bRedraw = True Then
            Call AddText("AutoRedraw = -1")
        Else
            Call AddText("AutoRedraw = 0")
        End If
        bFontTrans = ((bFlag And &H2) <> 0)
        If bFontTrans = True Then
            Call AddText("FontTransparent = -1")
        Else
            Call AddText("FontTransparent = 0")
        End If
        GetByte2 (F)
        ProccessPictureBox = 1
        Exit Function
    End If
    
    If bCode = 28 Then
        Call AddText("DrawStyle = " & GetByte2(F))
        ProccessPictureBox = 1
        Exit Function
    End If
    If bCode = 32 Then
        Call AddText("DrawMode = " & GetByte2(F))
        ProccessPictureBox = 1
        Exit Function
    End If
    If bCode = 50 Then
        Call AddText("Align = " & GetByte2(F))
        ProccessPictureBox = 1
        Exit Function
    End If
    If bCode = 42 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessPictureBox = 1
        Exit Function
    End If
    'I'f bCode = 63 Then
       ' Call AddText("OLEDragMode = " & GetByte2(f))
        ''ProccessPictureBox = 1
        'Exit Function
    'End If
    If bCode = 52 Then
        Dim dataLength As Integer
        dataLength = GetInteger(F)
        Call AddText("DataSource = " & cQuote & GetAllString(F) & cQuote)
        Exit Function
    End If
    
    If bCode = 66 Then 'DataFormat
        Call modControls.GetDataFormat(F, bEditor)
        ProccessPictureBox = 1
        Exit Function
    End If
    'Tabstop
    If bCode = 45 Then
        GetByte2 (F)
        Call AddText("TabStop = 0")
        ProccessPictureBox = 1
        Exit Function
    End If
    'Negotiate
    If bCode = 56 Then
        GetByte2 (F)
        Call AddText("Negotiate = -1")
        ProccessPictureBox = 1
        Exit Function
    End If
    'OLEDragMode
    If bCode = 63 Then
        Call AddText("OLEDragMode = " & GetByte2(F))
        ProccessPictureBox = 1
        Exit Function
    End If
    'OLEDropMode
    If bCode = 64 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessPictureBox = 1
        Exit Function
    End If
    'HasDC
    If bCode = 68 Then
    
        Call AddText("HasDC = " & GetByte2(F))
        ProccessPictureBox = 1
        Exit Function
    End If
    
    If bCode = 96 Then
        GetByte2 (F)
        ProccessPictureBox = 1
        Exit Function
    End If
    If bCode = 98 Then
        GetByte2 (F)
        ProccessPictureBox = 1
        Exit Function
    End If
    If bCode = 67 Then
        GetByte2 (F)
        ProccessPictureBox = 1
        Exit Function
    End If
    If bCode = 0 Then 'Null point
        GetByte2 (F)
        ProccessPictureBox = 1
        Exit Function
    End If
    If bCode = 99 Then 'Null  byte point
        GetByte2 (F)
        ProccessPictureBox = 1
        Exit Function
    End If
    ProccessPictureBox = -1
End Function


Public Function ProccessShape(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 4 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessShape = 8
        Exit Function
    End If
    If bCode = 11 Then
        Call AddText("Shape = " & GetByte2(F))
        ProccessShape = 1
        Exit Function
    End If
    If bCode = 12 Then
        Call AddText("DrawMode = " & GetByte2(F))
        ProccessShape = 1
        Exit Function
    End If
    If bCode = 13 Then
        Call AddText("BorderStyle = " & GetByte2(F))
        ProccessShape = 1
        Exit Function
    End If
    If bCode = 16 Then
        Call AddText("BackStyle = " & GetByte2(F))
        ProccessShape = 1
        Exit Function
    End If
    If bCode = 17 Then
        Call AddText("FillStyle = " & GetByte2(F))
        ProccessShape = 1
        Exit Function
    End If
    ProccessShape = -1
End Function
Public Function ProccessTextBox(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 4 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessTextBox = 8
        Exit Function
    End If
    If bCode = 10 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessTextBox = 1
        Exit Function
    End If
    If bCode = 19 Then
        Call AddText("BorderStyle = " & GetByte2(F))
        ProccessTextBox = 1
        Exit Function
    End If
    If bCode = 24 Then
        Call AddText("ScrollBars = " & GetByte2(F))
        ProccessTextBox = 1
        Exit Function
    End If

    If bCode = 36 Then
        Call AddText("Alignment = " & GetByte2(F))
        ProccessTextBox = 1
        Exit Function
    End If
    'DataSource
    If bCode = 41 Then
        Dim dataLength As Integer
        dataLength = GetInteger(F)
        Call AddText("DataSource = " & cQuote & GetAllString(F) & cQuote)
        ProccessTextBox = 2
        Exit Function
    End If
    If bCode = 46 Then
        Call frmMain.GetFontProperty(F)
        ProccessTextBox = 11
        Exit Function
    End If
    If bCode = 48 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessTextBox = 1
        Exit Function
    End If
    If bCode = 40 Then
        Call AddText("IMEMode=" & GetByte2(F))
        ProccessTextBox = 1
        Exit Function
    End If
    If bCode = 29 Then
        Call AddText("DragMode =" & GetByte2(F))
        ProccessTextBox = 1
        Exit Function
    End If
    If bCode = 53 Then
       Call AddText("OLEDropMode = " & GetByte2(F))
       ProccessTextBox = 1
       Exit Function
    End If
    If bCode = 52 Then
        Call AddText("OLEDragMode = " & GetByte2(F))
        ProccessTextBox = 1
        Exit Function
    End If
    
    If bCode = 56 Then 'DataFormat
        Call modControls.GetDataFormat(F, bEditor)
        ProccessTextBox = 1
        Exit Function
    End If

    If bCode = 0 Then
        Call GetByte2(F)
        ProccessTextBox = 1
        Exit Function
    End If
    ProccessTextBox = -1
End Function

Public Function ProccessTimer(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 7 Then
        AddText "Left=" & GetLong(F)
        ProccessTimer = 4
        Exit Function
    End If
    If bCode = 8 Then
        AddText "Top=" & GetLong(F)
        ProccessTimer = 4
        Exit Function
    End If
    ProccessTimer = -1
End Function
Public Function ProccessUserControl(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 53 Then
        Call AddText("ClientLeft = " & GetLong(F))
        Call AddText("ClientTop= " & GetLong(F))
        Call AddText("ClientWidth = " & GetLong(F))
        Call AddText("ClientHeight = " & GetLong(F))
        ProccessUserControl = 16
        Exit Function
    End If
    If bCode = 66 Then
        GetByte2 (F)
        ProccessUserControl = 1
        Exit Function
    End If
    If bCode = 64 Then
        Call frmMain.GetFontProperty(F)
        ProccessUserControl = 11
        Exit Function
    End If
    If bCode = 11 Then
        Call AddText("ClipBehavior = " & GetByte2(F))
        ProccessUserControl = 1
        Exit Function
    End If
    If bCode = 29 Then
        Call AddText("FillStyle = " & GetByte2(F))
        ProccessUserControl = 1
        Exit Function
    End If
    If bCode = 31 Then
        Call AddText("DrawMode = " & GetByte2(F))
        ProccessUserControl = 1
        Exit Function
    End If
    If bCode = 74 Then
        Call AddText("HitBehavior = " & GetByte2(F))
        ProccessUserControl = 1
        Exit Function
    End If
    If bCode = 87 Then
        Call AddText("BackStyle = " & GetByte2(F))
        ProccessUserControl = 1
        Exit Function
    End If
    If bCode = 89 Then
        Call AddText("BorderStyle = " & GetByte2(F))
        ProccessUserControl = 1
        Exit Function
    End If
    If bCode = 71 Then
        Call AddText("OLEDropMode = " & GetByte2(F))
        ProccessUserControl = 1
        Exit Function
    End If
    If bCode = 63 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessUserControl = 1
        Exit Function
    End If
    If bCode = 65 Then
        Call AddText("Appearance = " & GetByte2(F))
        ProccessUserControl = 1
        Exit Function
    End If
    ProccessUserControl = -1
End Function
Public Function ProccessVscroll(ByVal F As Variant, ByVal bCode As Byte, Optional ByVal bEditor As Boolean = False) As Integer
    If bCode = 2 Then
        If bEditor = False Then
            Call frmMain.GetControlSize(F)
        Else
            Call GetControlSizeEdit(F)
        End If
        ProccessVscroll = 8
        Exit Function
    End If
    If bCode = 8 Then
        Call AddText("MousePointer = " & GetByte2(F))
        ProccessVscroll = 1
        Exit Function
    End If
    If bCode = 16 Then
        Call AddText("DragMode = " & GetByte2(F))
        ProccessVscroll = 1
        Exit Function
    End If
    ProccessVscroll = -1
End Function

'##################################
'End Subs for Processing Special opcodes and properties for common controls
'##################################
Sub GetControlProperties(offset As Long)
'*****************************
'Purpose: Process Forms And Control Properties
'*****************************
Dim strCurrentForm As String
    'Erase existing data
        Dim fPos As Long 'Holds current location in the file used for controlheader
        Dim cListIndex As Integer ' Used for COM
        Dim cControlHeader As ControlHeader
        
 
        Dim FRXAddress As Long
        'Unload Old Controls
        If frmMain.txtEditArray.ubound > 0 Then
            For i = 1 To frmMain.txtEditArray.ubound
                Unload frmMain.txtEditArray(i)
                Unload frmMain.lblArrayEdit(i)
                Unload frmMain.cmdColor(i)
            Next
        End If

        
        Set gVBFile = New clsFile
        Call gVBFile.Setup(SFilePath)
        F = gVBFile.FileNumber
        Seek F, offset + 1
        FRXAddress = 0

        fPos = Loc(F)
        Dim posNextControl As Long
       
'Loop from new child control

        Seek F, fPos + 1
        Dim tArray As ArrayTestType
        Get F, , tArray
        Seek F, fPos + 1

        If tArray.arrayflag <> 128 Then
            Get #F, , cControlHeader
    
        Else
            Dim cArrayHeader As ControlArrayHeader
            Get #F, , cArrayHeader
           
            cControlHeader.Length = cArrayHeader.Length
            cControlHeader.cName = cArrayHeader.cName
            cControlHeader.cType = cArrayHeader.cType
            cControlHeader.cId = cArrayHeader.cId
            
        End If
        frmMain.lblObjectName.Caption = "ObjectName: " & cControlHeader.cName
        'With frmMain.pePropTree
        '    .LockWindowUpdate = True
        '    .Clear
        '     Set pMain = .PropertyItems.AddPropertyItem(cControlHeader.cName & " - Properties", "proptree")
        '        pMain.Expanded = True
        '    .Enabled = True
        '    .Visible = True
        '    .LockWindowUpdate = False
        'End With
        posNextControl = fPos + cControlHeader.Length + 2
        Dim fHeaderEnd As Long
        
        fHeaderEnd = Loc(F)

        Dim fControlEnd As Long
        'Store each object's information in a file
 
        Dim tliTypeInfo As TypeInfo 'Used for COM to find information about the properties of the control
        Dim FileLen As Long 'Used to caculate how much father to go in the control
        'Select what type of control it is
    
        Select Case cControlHeader.cType
            
            Case vbPictureBox '= 0
                cListIndex = 22
            Case vbLabel '= 1
                cListIndex = 14
       
            Case vbTextBox ' = 2
                cListIndex = 27
     
            Case vbFrame '= 3
                cListIndex = 10

            Case vbCommandbutton '= 4
                cListIndex = 4
    
            Case vbCheckbox '= 5
                cListIndex = 1
       
            Case vbOptionbutton     ' = 6
                cListIndex = 21
             
            Case vbComboBox     ' = 7
                cListIndex = 3
             
            Case vbListbox     '= 8
                cListIndex = 17
               
            Case vbHscroll     '= 9
                cListIndex = 12
            
            Case vbVscroll     '= 10
                cListIndex = 32
         
            Case vbTimer     '= 11
                cListIndex = 28
            Case vbform     '= 13
                cListIndex = 9
                strCurrentForm = cControlHeader.cName
            Case vbDriveListbox     '= 16
                cListIndex = 7
             
            Case vbDirectoryListbox     '= 17
                cListIndex = 6
        
            Case vbFileListbox     '= 18
                cListIndex = 8
            Case vbmenu     '= 19
                cListIndex = 19
            Case vbMDIForm     '= 20
                cListIndex = 18
                strCurrentForm = cControlHeader.cName
            Case vbShape     '= 22
                cListIndex = 26
            Case vbLine     '= 23
                cListIndex = 16
            Case vbImage     '= 24
                cListIndex = 13
            Case vbData     '= 37
                cListIndex = 5
            Case vbOLE     '= 38
                cListIndex = 20
            
            Case vbUserControl     '= 40
                cListIndex = 29
            
                strCurrentForm = cControlHeader.cName
  
            Case vbPropertyPage     '= 41
                cListIndex = 24
      
            Case vbUserDocument     '= 42
                cListIndex = 30
            Case 255 'external control
                'Load the control view COM if its on the computer
                'frmMain.pePropTree.Enabled = False
                Seek F, fPos + cControlHeader.Length ' - 2
                
        End Select
        Set tliTypeInfo = tliTypeLibInfo.GetTypeInfo(Replace(Replace(frmMain.lstTypeInfos.List(cListIndex), "<", ""), ">", ""))
        'Use the ItemData in lstTypeInfos to set the SearchData for lstMembers
        tliTypeLibInfo.GetMembersDirect frmMain.lstTypeInfos.ItemData(cListIndex), frmMain.lstMembers.hwnd, , , True
        
        FileLen = Loc(F) - fPos
        FileLen = cControlHeader.Length - FileLen
        
        Dim bCode As Byte 'holds gui opcode
        Dim varHold As Variant 'Holds the different data types
        Dim strHold As String 'holds the string
        Dim strReturnType As String 'holds the return type
      
        Do While Loc(F) < (fPos + cControlHeader.Length - 2)
       
            Get #F, , bCode
        
         FileLen = FileLen - 1
        
         Dim g As Long
         For g = 0 To frmMain.lstMembers.ListCount - 1
            
    'Process special events
            Dim iFileLength As Integer
        Select Case cControlHeader.cType
                Case vbPictureBox: '0
                    iFileLength = modControls.ProccessPictureBox(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbLabel: ' = 1
                    iFileLength = modControls.ProccessLabel(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbTextBox: ' = 2
                    iFileLength = modControls.ProccessTextBox(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbFrame: ' = 3
                     iFileLength = modControls.ProccessFrame(F, bCode, True)
                     If iFileLength <> -1 Then Exit For
                Case vbCommandbutton: ' = 4
                     iFileLength = modControls.ProccessCommandButton(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbCheckbox: '= 5
                    iFileLength = modControls.ProccessCheckBox(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbOptionbutton: ' = 6
                    iFileLength = modControls.ProccessOption(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbComboBox: ' = 7
                    iFileLength = modControls.ProccessComboBox(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbListbox: ' = 8
                    iFileLength = modControls.ProccessListBox(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbHscroll: ' = 9
                    iFileLength = modControls.ProccessHscroll(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbVscroll: ' = 10
                    iFileLength = modControls.ProccessVscroll(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbTimer: ' = 11
                    iFileLength = modControls.ProccessTimer(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbform: ' = 13
                    iFileLength = modControls.ProccessForm(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbDriveListbox: ' = 16
                    iFileLength = modControls.ProccessDriveListBox(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbDirectoryListbox: ' = 17
                    iFileLength = modControls.ProccessDirListBox(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbFileListbox: ' = 18
                    iFileLength = modControls.ProccessFileListBox(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbmenu: ' = 19
                    iFileLength = modControls.ProccessMenu(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbMDIForm: ' = 20
                    iFileLength = modControls.ProccessForm(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbShape: ' = 22
                    iFileLength = modControls.ProccessShape(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbLine: ' = 23
                   iFileLength = modControls.ProccessLine(F, bCode, True)
                   If iFileLength <> -1 Then Exit For
                Case vbImage: ' = 24
                    iFileLength = modControls.ProccessImage(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbData: ' = 37
                    iFileLength = modControls.ProccessDataControl(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                    
                Case vbOLE: ' = 38
                    iFileLength = modControls.ProccessOLE(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbUserControl: ' = 40
                    
                    iFileLength = modControls.ProccessUserControl(F, bCode, True)
                    If iFileLength <> -1 Then Exit For
                Case vbPropertyPage: ' = 41
                Case vbUserDocument: ' = 42
                
            End Select
            
            strReturnType = "n"
            If ReturnGuiOpcode(frmMain.lstTypeInfos.ItemData(cListIndex), tliInvokeKinds, frmMain.lstMembers.List(g)) = bCode Then
              Dim strHelp As String
                strReturnType = Trim$(ReturnDataType(frmMain.lstTypeInfos.ItemData(cListIndex), tliInvokeKinds, frmMain.lstMembers.List(g)))
                strHelp = modCOM.ReturnHelpString(frmMain.lstTypeInfos.ItemData(cListIndex), tliInvokeKinds, frmMain.lstMembers.List(g))
        

                If InStr(1, strReturnType, "Byte") Then
                    varHold = GetByte2(F)
                    Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), varHold, "Byte", Loc(F) - 1, strHelp)
                 
                    FileLen = FileLen - 1
                    Exit For
                End If
                If InStr(1, strReturnType, "Boolean") Then
                    varHold = GetBoolean(F)
                    If varHold = True Then
                         Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), "False", "Boolean", Loc(F) - 2, strHelp)
               
                    Else
                         Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), "True", "Boolean", Loc(F) - 2, strHelp)
               
                    End If
                    Seek F, Loc(F)
                    FileLen = FileLen - 2
                    Exit For
                End If
                If InStr(1, strReturnType, "Integer") Then
                    varHold = gVBFile.GetInteger(Loc(F))
                    Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), varHold, "Integer", Loc(F) - 2, strHelp)
                   
                    FileLen = FileLen - 2
                    Exit For
                End If
                If InStr(1, strReturnType, "Long") Then
                    varHold = GetLong(F)
                    Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), varHold, "Long", Loc(F) - 4, strHelp)
                 
                    FileLen = FileLen - 4
                    Exit For
                End If
                
                If InStr(1, strReturnType, "Single") Then
                    varHold = GetSingle(F)
                    Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), varHold, "Single", Loc(F) - 4, strHelp)
                    'Call AddText(lstMembers.List(g) & " = " & varHold & strExtraInfo)
                    FileLen = FileLen - 4
                    Exit For
                End If

                If InStr(1, strReturnType, "String") Then
          
                    strHold = GetAllString(F)
                    Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), strHold, "String", Loc(F), strHelp)
                    'Call AddText(lstMembers.List(g) & " = " & cQuote & strHold & cQuote & strExtraInfo)
                    FileLen = FileLen - Len(strHold) - 3
                    Exit For
                End If

              
                If InStr(1, strReturnType, "stdole.Picture") Then
                    
                    varHold = GetLong(F)
                   
                    If varHold <> -1 Then
                  
                        If cControlHeader.cName <> strCurrentForm Then
                            Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), varHold, "Picture", Loc(F), strHelp)
                            Call frmMain.GetStdPicture(F, varHold, strCurrentForm & "." & cControlHeader.cName, strCurrentForm, FRXAddress)
                            
                        Else
                            Call modGlobals.AddPropertyToTheList(frmMain.lstMembers.List(g), varHold, "Picture", Loc(F), strHelp)
                            Call frmMain.GetStdPicture(F, varHold, cControlHeader.cName, strCurrentForm, FRXAddress)
                          
                        End If
                        
                        
    
                        Seek F, Loc(F)
                     
                       FRXAddress = FRXAddress + varHold + 12
               
             
                        FileLen = FileLen - 12
        
                    Else
                        FileLen = FileLen - 4
                    End If
                    Exit For
                End If
               
               
                Exit For

            End If
            

         Next
         'Check For Unknown opcode
         If strReturnType = "n" Then
            'UNKOWN Opcode
            Call modGlobals.AddToErrorLog("Unknown OpCode Control Edit: " & cControlHeader.cName & " Bcode:" & bCode & " CType: " & cControlHeader.cType & " LOC: " & Loc(F))
            Seek F, fPos + cControlHeader.Length
            Exit Do
         End If
         
         'Exit the Process controls in case it hangs on a property
         If CancelDecompile = True Then Exit Sub
         DoEvents
        Loop
        
        
        'Get the seperator type for the end of the control
        Dim cControlEnd As Integer
        cControlEnd = GetInteger(F)


        If cControlEnd = vbFormNewChildControl Then
            Seek F, posNextControl
            
            
        End If
        
        If cControlEnd = vbFormExistingChildControl Then
            Dim bCheckEnd As Byte
            
            Do
                Get F, , bCheckEnd
            Loop Until bCheckEnd = 3 Or bCheckEnd = 4 Or bCheckEnd >= 5
        End If

'##########################################
'End of Form/Control Properties Loop
'##########################################
End Sub


Sub GetControlSizeEdit(F As Variant)
'*****************************
'Purpose: Get the control size type
'*****************************
    Dim cPosition As typeStandardControlSize
    Dim fPos As Long
    fPos = Loc(F) + 1
    Get F, , cPosition
    If cPosition.cLeft <> -32768 Then
    
        Call modGlobals.AddPropertyToTheList("Left", cPosition.cLeft, "Integer", Loc(F) - 8, "Returns/sets the distance between the internal left edge of an object and the left edge of its container.")
        Call modGlobals.AddPropertyToTheList("Top", cPosition.cTop, "Integer", Loc(F) - 6, "Returns/sets the distance between the internal top edge of an object and the top edge of its container.")
        Call modGlobals.AddPropertyToTheList("Height", cPosition.cHeight, "Integer", Loc(F) - 4, "Returns/sets the height of an object.")
        Call modGlobals.AddPropertyToTheList("Width", cPosition.cWidth, "Integer", Loc(F) - 2, "Returns/sets the width of an object.")
           
    Else
     Dim cPosition2 As typeStandardControlSize2
     Get F, fPos + 2, cPosition2
        Call modGlobals.AddPropertyToTheList("Left", cPosition2.cLeft, "Long", Loc(F) - 16, "Returns/sets the distance between the internal left edge of an object and the left edge of its container.")
        Call modGlobals.AddPropertyToTheList("Top", cPosition2.cTop, "Long", Loc(F) - 12, "Returns/sets the distance between the internal top edge of an object and the top edge of its container.")
        Call modGlobals.AddPropertyToTheList("Height", cPosition2.cHeight, "Long", Loc(F) - 8, "Returns/sets the height of an object.")
        Call modGlobals.AddPropertyToTheList("Width", cPosition2.cWidth, "Long", Loc(F) - 4, "Returns/sets the width of an object.")

    End If

End Sub

