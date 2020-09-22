VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniRW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ini RW Class
'Created by: Fida Shah
'Date: September 11th 2000
'Submitted to: Planet source code

Option Explicit

'***********************************************************************************
'***********************************************************************************
'Declaration for read/write from/to ini files
#If Win32 Then
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
    lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
    lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
    Private Declare Function WritePrivateProfileString Lib "Kernel" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpString As Any, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileString Lib "Kernel" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpDefault As Any, ByVal lpReturnedString As String, _
    ByVal nSize As Integer, ByVal lpFileName As String) As Integer
#End If

'Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long

'Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, _
ByVal nSize As Long, ByVal lpFileName As String) As Long
'End declaration for read/write from/to ini files
'***********************************************************************************
'***********************************************************************************



'***********************************************************************************
'***********************************************************************************
'Declaration of private variables
Private Filename As String
Private fileExist As Boolean
'End declaration of private variables
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Property Let File_Name(fName As String)
    'First check if file exists. We need to do that for all operations other than
    'adding sections/keys/values. So we call the function CheckExistence by
    'passing in the file name given by the user. That function will first check
    'if the file exists or not. If the file does not exist, it will attempt
    'to create a file of that name. If the file could not be created, an empty
    'string will be returned, otherwise the initial file name will be returned.
    'The return value of the function will be assigned to fileName variable
    'of the class
    
    'Set the name/path of the file
    Filename = CheckExistence(fName)
    
End Property
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Property Get File() As String
    'Get the name/path of the file
    File = Filename
End Property
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Property Get FileOpen() As Boolean
    'Get the boolean value for file existence
    FileOpen = fileExist
End Property
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Private Function CheckExistence(fName As String) As String
'note for myself: working
    'This function will first check if the file exists or not.
    'If the file does not exist, it will attempt to create an empty ini
    'file by the file name and path given by the user. If the file could
    'not be created, an empty string will be returned, otherwise the
    'initial file name will be returned.
    
    'Set error handler
    On Error GoTo LocalErrHandler
    
    'Set default values value
    CheckExistence = fName
    fileExist = True
    
    'If file does not exist
    If Dir$(Filename) <> "" Then
        'We will attempt to create an empty ini file of the given name and path
        
        'Write to file, if file not created
        If IIf(WritePrivateProfileString("", vbNullString, vbNullString, fName), False, True) Then
            'Set return value
            CheckExistence = ""
            'Set value of fileExist
            fileExist = False
        End If
    End If
    
    'Exit Function before going to error handler
    Exit Function
    
LocalErrHandler:
    'Call error handler function
    ErrorHandler Err.Number, Err.Description, ""
End Function
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Public Function WriteData(SectionName As String, keyName As String, keyValue As String) As Boolean
'note for myself: working
    'This function writes data to ini file and returns true or false
    'according to operation success
    'Following operations can be performed with this function
    '   -Create an empty section
    '   -Create a section with an empty key
    '   -Create a section with a key having some value
    '   -Change value of an existing key
    
    'Set error handler
    On Error GoTo LocalErrHandler
    
    'If file does not exist or could not be created then exit function
    If Not fileExist Then Exit Function
    
    'Write to file, if successful operation made then return true
    WriteData = IIf(WritePrivateProfileString(SectionName, keyName, _
            keyValue, Filename), True, False)
    
    'Exit Function before going to error handler
    Exit Function
    
LocalErrHandler:
    'Call error handler function
    ErrorHandler Err.Number, Err.Description, ""
End Function
'***********************************************************************************
'***********************************************************************************


'***********************************************************************************
'***********************************************************************************
Public Function DeleteData(SectionName As String, keyName As String, keyValue) As Boolean
'note for myself: working
    'This function deletes data from ini file and returns true or false
    'according to operation success
    'Following operations can be performed with this function
    '   -Delete a key
    '   -Delete an entire section
    'To just delete the value, use WriteData function and pass in "" as key value
    'However, for this and above operations the specified key and section should
    'exist
    
    'Set error handler
    On Error GoTo LocalErrHandler
    
    'If file does not exist or could not be created then exit function
    If Not fileExist Then Exit Function
    
    'Write to file, if successful operation made then return true, other wise false
    DeleteData = IIf(WritePrivateProfileString(SectionName, keyName, _
                keyValue, Filename), True, False)
    
    'Exit Function before going to error handler
    Exit Function
    
LocalErrHandler:
    'Call error handler function
    ErrorHandler Err.Number, Err.Description, ""
End Function
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Public Function GetAllSections() As Variant
'note for myself: working
    'This function retrievs ALL SECTIONS (just names) in the file
    'and returns them as an variant array
    
    'Set error handler
    On Error GoTo LocalErrHandler
    
    'If file does not exist or could not be created then exit function
    If Not fileExist Then Exit Function
    
    'Declare local variables
    Dim loopCounter As Integer          'Loop counter
    Dim arrayCounter As Integer         'Array counter
    Dim strData As String               'String containg all sections
    Dim ret As Long                     'Value returned by API
    Dim sectionsToReturn As Variant     'Array to store all the sections retrieved
    
    'set string length
    'You may want to increase this size for huge ini files
    strData = String$(255, 0)
    
    'Initialize array counter
    arrayCounter = 0
        
    'Allocate the very first element for the array
    ReDim sectionsToReturn(arrayCounter) As Variant
    
    'Add an empty string to the 0-element of the array. This is in case
    'no sections were found in the file
    sectionsToReturn(0) = ""
    
    'This call will put all the section names in the file in "strData" separated by
    'chr$(0). "ret" will indicate number of characters in "strData" including
    'chr$(0)
    ret = GetPrivateProfileString(vbNullString, vbNullString, _
            vbNullString, strData, Len(strData), Filename)
    
    'The returned value indicates number of characters in the string
    'So if 'Ret' > 0 then atleast one key is returned
    If (ret > 0) Then
        'Loop through the string to separate and distribute the section names
        'this is being done by adding character by character from "strData" to
        'the array element until Chr(0) is encountered. The we jump to next array
        'element
        For loopCounter = 1 To ret - 1
            'Compare each character with Chr(0)
            If StrComp(Mid(strData, loopCounter, 1), Chr(0), vbBinaryCompare) = 0 Then
                'Increment array counter
                arrayCounter = arrayCounter + 1
                
                'If the character is Chr(0), then allocate a new array element
                ReDim Preserve sectionsToReturn(arrayCounter) As Variant
            Else
                'Add each character to the array element
                sectionsToReturn(arrayCounter) = sectionsToReturn(arrayCounter) & Mid(strData, loopCounter, 1)
            End If
        Next
    End If
    
    'Return array
    GetAllSections = sectionsToReturn
    
    'Exit Function before going to error handler
    Exit Function
    
LocalErrHandler:
    'Call error handler function
    ErrorHandler Err.Number, Err.Description, ""
End Function
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Public Function GetAllKeysInSection(SectionName As String) As Variant
'note for myself: working
    'This function retrievs ALL KEYS (excluding their values) in the section
    'and returns them as an variant array
    
    'Set error handler
    On Error GoTo LocalErrHandler
    
    'If file does not exist or could not be created then exit function
    If Not fileExist Then Exit Function
    
    'Declare local variables
    Dim loopCounter As Integer      'Loop counter
    Dim arrayCounter As Integer     'Array counter
    Dim strData As String           'String containg all keynames
    Dim ret As Long                 'Value returned by API
    Dim keysToReturn() As Variant   'Array to store all the keys retrieved
    
    'set string length
    'You may want to increase this size for huge ini files
    strData = String$(30000, 0)
    
    'Initialize array counter
    arrayCounter = 0
        
    'Allocate the very first element for the array
    ReDim keysToReturn(arrayCounter) As Variant
    
    'Add an empty string to the 0-element of the array. This is in case
    'no keys were found in the section
    keysToReturn(0) = ""
    
    'This call will put all the keynames in the section in "strData" separated by
    'chr$(0). "ret" will indicate number of characters in "strData" including
    'chr$(0)
    ret = GetPrivateProfileString(SectionName, vbNullString, _
            vbNullString, strData, Len(strData), Filename)
    
    'The returned value indicates number of characters in the string
    'So if 'Ret' > 0 then atleast one key is returned
    If (ret > 0) Then
        'Loop through the string to separate and distribute the key names
        'this is being done by adding character by character from "strData" to
        'the array element until Chr(0) is encountered. The we jump to next array
        'element
        For loopCounter = 1 To ret - 1
            'Compare each character with Chr(0)
            If StrComp(Mid(strData, loopCounter, 1), Chr(0), vbBinaryCompare) = 0 Then
                'Increment array counter
                arrayCounter = arrayCounter + 1
                
                'If the character is Chr(0), then allocate a new array element
                ReDim Preserve keysToReturn(arrayCounter) As Variant
            Else
                'Add each character to the array element
                keysToReturn(arrayCounter) = keysToReturn(arrayCounter) & Mid(strData, loopCounter, 1)
            End If
        Next
    End If
    
    'Return array
    GetAllKeysInSection = keysToReturn
    
    'Exit Function before going to error handler
    Exit Function
    
LocalErrHandler:
    'Call error handler function
    ErrorHandler Err.Number, Err.Description, ""
End Function
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Public Function GetAllKeysValuesInSection(SectionName As String) As Variant
'note for myself: working
    'This function retrievs ALL KEYS AND THEIR VALUES in the section
    'and returns them as a 2-dimensional variant array
    
    'Set error handler
    On Error GoTo LocalErrHandler
    
    'If file does not exist or could not be created then exit function
    If Not fileExist Then Exit Function
    
    'Declare local variables
    Dim counter As Integer          'Loop counter
    Dim keyValue As String          'Store retrieved value here
    Dim allKeys As Variant          'All the keys in section
    Dim allKeysValues As Variant  'All the keys and their values in section
    
    'Set string length
    keyValue = String$(255, 0)
    
    'Call function to get all keys in the section
    allKeys = GetAllKeysInSection(SectionName)
    
    'Allocate the very first element for the array
    ReDim allKeysValues(UBound(allKeys), 1) As Variant
    
    'Loop through allkeys-array
    For counter = LBound(allKeys) To UBound(allKeys)
        'for each key in allkeys-array, copy the key into the
        '1-st dimension of allKeysValues-array and that key's
        'value into the 2-nd dimension of allKeysValues-array
        allKeysValues(counter, 0) = allKeys(counter)
        allKeysValues(counter, 1) = IIf(GetPrivateProfileString(SectionName, _
                                    CStr(allKeys(counter)), "", keyValue, _
                                    Len(keyValue), Filename), Left(keyValue, _
                                    InStr(keyValue, Chr(0)) - 1), "")
            
    Next
    
    'return this 2-dimensional array
    GetAllKeysValuesInSection = allKeysValues
    
    'Exit Function before going to error handler
    Exit Function
    
LocalErrHandler:
    'Call error handler function
    ErrorHandler Err.Number, Err.Description, ""
End Function
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Public Function GetValue(SectionName As String, keyName As String) As String
'note for myself: working
    'This function returns value of a specified key
    
    'Set error handler
    On Error GoTo LocalErrHandler
    
    'If file does not exist or could not be created then exit function
    If Not fileExist Then Exit Function
    
    'Declare local variables
    Dim keyValue As String  'Store retrieved value here
    
    'Set string length
    keyValue = String$(255, 0)
        
    'Read from file and section. If key is found then return its value,
    'otherwise return empty string
    GetValue = IIf(GetPrivateProfileString(SectionName, keyName, _
            "", keyValue, Len(keyValue), Filename), Left(keyValue, _
            InStr(keyValue, Chr(0)) - 1), "")
    
    'Exit Function before going to error handler
    Exit Function
    
LocalErrHandler:
    'Call error handler function
    ErrorHandler Err.Number, Err.Description, ""
End Function
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Public Function FindKeyInFile(keyName As String) As Variant
'note for myself: workinging on it
    'This Function searches the entire file to find a key. It returns
    'an array with the section names contianing the key
    
    'Set error handler
    On Error GoTo LocalErrHandler
    
    'If file does not exist or could not be created then exit function
    If Not fileExist Then Exit Function
    
    'Declare local variables
    Dim counter As Integer          'Loop counter
    Dim arrayCounter As Integer     'Array counter
    Dim allSections As Variant      'All sections in file
    Dim sectionsWithKeys As Variant 'All sections containing the key
    
    'Initialize array counter
    arrayCounter = 0
    
    'Allocate the very first element for the arrays
    ReDim sectionsWithKeys(arrayCounter) As Variant
    
    'Add an empty string to the 0-element of the array. This is in case
    'no keys were found in the file
    sectionsWithKeys(0) = ""

    'Call function to get all sections in the file
    allSections = GetAllSections
    
    'Check if there was any section in the file
    If UBound(allSections) > -1 And CStr(allSections(0)) <> "" Then
        'Loop thru sections to find key
        For counter = LBound(allSections) To UBound(allSections)
            'Call KeyExistInSection-function to see if key exists or not
            If FindKeyInSection(CStr(allSections(counter)), keyName) Then
                'If key exists, then add en element to the sectionsWithKeys-array
                ReDim Preserve sectionsWithKeys(arrayCounter) As Variant
                'Add the section name in that element
                sectionsWithKeys(arrayCounter) = allSections(counter)
                'increment array counter
                arrayCounter = arrayCounter + 1
            End If
        Next
    End If
    
    'Return array
    FindKeyInFile = sectionsWithKeys
    
    'Exit Function before going to error handler
    Exit Function

LocalErrHandler:
    'call error handler function
    ErrorHandler Err.Number, Err.Description, ""
End Function
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Public Function FindKeyInSection(SectionName As String, keyName As String) As Boolean
'note for myself: working
    'This Function searches a specified section to find a key. It returns
    'true if match found otherwise it reutrns false
    
    'Set error handler
    On Error GoTo LocalErrHandler
    
    'If file does not exist or could not be created then exit function
    If Not fileExist Then Exit Function
    
    'Declare local variables
    Dim counter As Integer          'Loop counter
    Dim keysInSection As Variant    'All the keys in section
    
    'Set default return
    FindKeyInSection = False

    'Call function to get all keys in the section
    keysInSection = GetAllKeysInSection(SectionName)
    
    'Loop through all keys to find match
    For counter = LBound(keysInSection) To UBound(keysInSection)
        'If match found, set return value to true and exit function
        If StrComp(CStr(keysInSection(counter)), keyName, 1) = 0 Then
            FindKeyInSection = True
            Exit Function
        End If
    Next
        
    'Exit Function before going to error handler
    Exit Function

LocalErrHandler:
    'call error handler function
    ErrorHandler Err.Number, Err.Description, ""
End Function
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Public Function FindValueInSection(SectionName As String, keyValue As String) As Variant
'note for myself: working
    'This Function searches a specified section to find a value. It returns
    'an array witht the key names containing that value
    
    'Set error handler
    On Error GoTo LocalErrHandler
    
    'If file does not exist or could not be created then exit function
    If Not fileExist Then Exit Function
    
    'Declare local variables
    Dim counter As Integer              'Loop counter
    Dim arrayCounter As Integer         'Array counter
    Dim keysValuesInSection As Variant  'All the keys and values in section
    Dim keysToReturn As Variant         'Keys containing the value

    'Initialize array counter
    arrayCounter = 0
    
    'Allocate the very first element for the arrays
    ReDim keysToReturn(arrayCounter) As Variant
    
    'Add an empty string to the 0-element of the array. This is in case
    'no values were found in the section
    keysToReturn(0) = ""
    
    'Call function to get all keys and values in the section
    keysValuesInSection = GetAllKeysValuesInSection(SectionName)
        
    'Loop through all keys to find match
    For counter = LBound(keysValuesInSection) To UBound(keysValuesInSection)
        'If match found
        If StrComp(CStr(keysValuesInSection(counter, 1)), keyValue, 1) = 0 Then
            'Add an element to keysToReturn-array
            ReDim Preserve keysToReturn(arrayCounter) As Variant
            'Add key name to the element
            keysToReturn(arrayCounter) = keysValuesInSection(counter, 0)
            'Increment array counter
            arrayCounter = arrayCounter + 1
        End If
    Next
    
    'Return array
    FindValueInSection = keysToReturn
        
    'Exit Function before going to error handler
    Exit Function

LocalErrHandler:
    'call error handler function
    ErrorHandler Err.Number, Err.Description, ""
End Function
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Public Function FindValueInFile(keyValue As String) As Variant
'note for myself: working
    'This Function searches for a specified value in the file. It returns
    'an array witht the section and key names containing that value
    
    'Set error handler
    'On Error GoTo LocalErrHandler
    
    'If file does not exist or could not be created then exit function
    If Not fileExist Then Exit Function
    
    'Declare local variables
    Dim counter1 As Integer             'Loop counter
    Dim counter2 As Integer             'Loop counter
    Dim arrayCounter As Integer         'Array counter
    Dim temp1 As Variant                'Temporary array
    Dim temp2 As Variant                'Temporary array
    Dim allSections As Variant          'All sections in file
    Dim keysInSection As Variant        'All the keys and values in section
    Dim sectionsKeysToReturn As Variant 'Keys containing the value

    'Initialize array counter
    arrayCounter = 0
    
    'Allocate the very first element for the array
    ReDim temp1(arrayCounter) As Variant
    ReDim temp2(arrayCounter) As Variant
    
    'Put an empty string the element
    temp1(arrayCounter) = ""
    temp2(arrayCounter) = ""
    
    'Call function to get all sections in file
    allSections = GetAllSections
        
    'Loop through all sections
    For counter1 = LBound(allSections) To UBound(allSections)
        'For each section get all list of keys containg the value
        'Simply call function FindValueInSection
        keysInSection = FindValueInSection(CStr(allSections(counter1)), keyValue)
        
        'Check if any keys were returned at all
        If UBound(keysInSection) > -1 And keysInSection(0) <> "" Then
            'if key were returned then add elements to the
            'temp-array
            ReDim Preserve temp1(arrayCounter + UBound(keysInSection)) As Variant
            ReDim Preserve temp2(arrayCounter + UBound(keysInSection)) As Variant
            
            'Loop to add returned keys  and the current section in the array
            For counter2 = LBound(keysInSection) To UBound(keysInSection)
                'Add section name
                temp1(arrayCounter) = allSections(counter1)
                'Add key name
                temp2(arrayCounter) = keysInSection(counter2)
                'Increment array counter
                arrayCounter = arrayCounter + 1
            Next
        End If
    Next
    
    'temp-array is now holding all the section and keynames we need
    'Allocate elements required for the array that will be returned
    ReDim sectionsKeysToReturn(UBound(temp1), 1) As Variant
    
    'Add an empty string to the 0-elements of the array. This is in case
    'no values were found in the section
    sectionsKeysToReturn(0, 0) = ""
    sectionsKeysToReturn(0, 1) = ""
    
    'Loop through temp-array
    For counter1 = LBound(temp1) To UBound(temp1)
        'Copy data from temp1 and temp2 to sectionsKeysToReturn-array
        sectionsKeysToReturn(counter1, 0) = temp1(counter1)
        sectionsKeysToReturn(counter1, 1) = temp2(counter1)
    Next
    
    'Return array
    FindValueInFile = sectionsKeysToReturn
        
    'Exit Function before going to error handler
    Exit Function

LocalErrHandler:
    'call error handler function
    ErrorHandler Err.Number, Err.Description, ""
End Function
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Public Function RenameSection(SectionName As String, newSectionName As String) As Boolean
'note for myself:
    
    'This function renames a section. It is done by first getting all the keys and
    'their values in the section and storing them in an array. Then the section is
    'deleted and a new section is created with the new name. Then all the keys and
    'values in the array are added to that new section
    
    'Set error handler
    On Error GoTo LocalErrHandler
    
    'If file does not exist or could not be created then exit function
    If Not fileExist Then Exit Function
    
    'Declare local variables
    Dim counter As Integer          'loop counter
    Dim allKeysValues As Variant    'values of keys
    Dim success As Boolean          'success flag
    
    'Set default return
    RenameSection = False
    
    'get all keys and their values in the section
    allKeysValues = GetAllKeysValuesInSection(SectionName)
    
    'Call function to delete the entire section
    success = DeleteData(SectionName, vbNullString, vbNullString)
    
    'If section successfully deleted then proceed
    If success Then
        'Start writing the keys and values in new section
        '(The section will be created automatically)
        
        'NOTE: if the section did not have any keys in it,
        'the allKeysValues-array will have only one element in it and that
        'element is going to contain an ampty string. If that is the case then
        'we are actually just creating an empty section (just section name),
        'which is done in two steps'
        'Step 1: create the section with a dummy key
        'Step 2: delete the dummy key
        
        'Check if the there is one one element returned in allKeysValues-array
        If UBound(allKeysValues) = 0 Then
            'Check if that element has any key in it
            If CStr(allKeysValues(0, 0)) = "" Then
                'There was no key in this element
            
                'First create the section with a dummy key
                If WriteData(newSectionName, "dummy", "") Then
                    'If section with dummy key created, then delete the dummy key
                    If WriteData(newSectionName, "dummy", vbNullString) Then
                        'Set new return value
                        RenameSection = True
                        
                        'Exit function
                        Exit Function
                    End If
                End If
            End If
        End If
        
        'loop to write keys to the section
        For counter = LBound(allKeysValues) To UBound(allKeysValues)
            WriteData newSectionName, CStr(allKeysValues(counter, 0)), _
                                        CStr(allKeysValues(counter, 1))
        Next
        
        'Set new return value
        RenameSection = True
    End If
        
    'Exit Function before going to error handler
    Exit Function
    
LocalErrHandler:
    'call error handler function
    ErrorHandler Err.Number, Err.Description, ""
End Function
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Public Function RenameKey(SectionName As String, keyName As String, newKeyName As String) As Boolean
'note for myself:
    'This function renames a key. It is done by strong the value of the given key,
    'then deleting the key, and at the end assiging the value to the new key
        
    'Set error handler
    On Error GoTo LocalErrHandler
    
    'If file does not exist or could not be created then exit function
    If Not fileExist Then Exit Function
    
    'Declare local variables
    Dim value As String    'value of the key
    Dim success As Boolean  'success flag
    
    'Set default return
    RenameKey = False
    
    'get the key value
    value = GetValue(SectionName, keyName)
    
    'Call function to delete the key
    success = DeleteData(SectionName, keyName, vbNullString)
    
    'If section successfully deleted then proceed
    If success Then
        'Write the new key and the old value
        WriteData SectionName, newKeyName, value
        
        
        'Set new return value
        RenameKey = True
    End If
        
    'Exit Function before going to error handler
    Exit Function
    
LocalErrHandler:
    'call error handler function
    ErrorHandler Err.Number, Err.Description, ""
End Function
'***********************************************************************************
'***********************************************************************************

'***********************************************************************************
'***********************************************************************************
Private Function ErrorHandler(ErrorNumber As Integer, ErrorDescription As String, AdditionalMessage As String)
    'if additional message was passed in
    If AdditionalMessage <> "" Then
        MsgBox "Error Number: " & ErrorNumber & vbCrLf & ErrorDescription & _
                vbCrLf & vbCrLf & AdditionalMessage, _
                vbCritical, "Ini RW Class"
    Else
        MsgBox "Error Number: " & ErrorNumber & vbCrLf & _
                ErrorDescription, vbCritical, "Ini RW Class"
    End If
End Function
'***********************************************************************************
'***********************************************************************************


