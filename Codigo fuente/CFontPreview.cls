VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFontPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Proyecto dx_lib32                                        
'-------------------------------------------------------------------------------
'                                                          
' Copyright (C) 2001 - 2010, José Miguel Sánchez Fernández 
'                                                          
' This file is part of dx_lib32 project.
'
' dx_lib32 project is free software: you can redistribute it and/or modify
' it under the terms of the GNU Lesser General Public License as published by
' the Free Software Foundation, version 2 of the License.
'
' dx_lib32 is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public License
' along with dx_lib32 project. If not, see <http://www.gnu.org/licenses/>.
'===============================================================================

'===============================================================================
' Name: DirectShowEvent
' Purpose: Modulo de sistema de importacion de fuentes True Type no instaladas
' Functions:
'     <functions' list in alphabetical order>
' Properties:
'     <properties' list in alphabetical order>
' Methods:
'     <Methods' list in alphabetical order>
' Author: José Miguel Sánchez Fernández
' Start: 01/12/2008
' Modified: 11/11/2009
'===============================================================================

' Original author: Karl E. Peterson
' Source: http://vb.mvps.org/articles/ap200107.asp

' ****************************************************************
'  Related KB, Q112672, to make installation permanent...
'  HOWTO: Add a Scalable Font to Windows from Visual Basic
'  http://support.microsoft.com/support/kb/articles/Q112/6/72.htm
'  Related KB, Q130459...
'  Adding TrueType, Raster, or Vector Fonts to System
'  http://support.microsoft.com/support/kb/articles/Q130/4/59.htm
'  Lots more stuff on typography:
'  http://www.microsoft.com/typography/default.htm
' ****************************************************************
'  Portions of this code inspired by the detective work of
'  of Mathias Schiffer <Schiffer@aixsoft.de>, 04/13/2001:
' ----------------------------------------------------------------
'  TTF specs can be found on
'  http://www.microsoft.com/typography/tt/tt.htm
'
'  I used "chapter 2 - the TrueType font file [0.3Mb Word]"
'  http://www.microsoft.com/typography/tt/ttf_spec/ttch02.doc
'  from that resource. Page numbers relate to that doc.
'
'  A TTF starts off with an offset table. What follows is a list
'  of TableDirectory entries (NumberOfTables in the offset table
'  tells you how many of them there are). Each TableDirectory
'  entry refers to a "NamingTable" that includes "NameRecords"
'  and storage space for the information we are looking for.
'
'  Notice: All of that stuff is unsinged data unless stated
'          otherwise. Shouldn't matter much, though. What matters
'          is it's stored in Big Endian byte format, so you need
'          to reverse bytes to get the actual info.
' ****************************************************************
Option Explicit

' Win32 APIs
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpszFileName As String) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const WM_FONTCHANGE As Long = &H1D

' TrueType Font file table structures...
' ****************************************************************
' The TrueType font file begins at byte 0 with the Offset Table
Private Type OFFSETTABLE ' (page 32: "The Table Directory")
   Version As Long           ' signed floating point number
                             ' 2.14 (0x00010000 for TTF version 1.0)
   NumberOfTables As Integer ' number of tables
   SearchRange As Integer    ' maximum power of 2 <= numTables) x 16
   EntrySelector As Integer  ' Log2: maximum power of 2 <= numTables)
   RangeShift As Integer     ' NumTables * (16 - searchRange)
End Type

' The offset table is followed at byte 12 by the Table
' Directory entries. Entries in the Table Directory are
' sorted in ascending order by tag.
Private Type TABLEDIRECTORY ' (page 32: "The Table Directory")
   Tag As String * 4 ' (4-byte identifier). Tags are the names given
                     ' to tables in the file (e.g. "name" for string info).
   Checksum As Long  ' CheckSum for this table - may come in handy to
                     ' ensure the file really is a TTF file
   Offset As Long    ' Offset of this table (from beginning of file)
   Length As Long    ' Length of this table
End Type

Private Type FONTHEADER ' (page 65: "head - Font Header")
   ' We're not interested in anything but the MagicNumber
   ' member here. Most other members have been
   ' set to have the right number of bits only.
   TableVersionNumber As Long ' FIXED
   FontRevision As Long       ' FIXED
   CheckSumAdjustment As Long
   MagicNumber As Long   ' Needs to be &H5F0F3CF5&
   Flags As Integer
   UnitsPerEm As Integer
   Created As Currency   ' longDateTime
   Modified As Currency  ' longDateTime
   xMin As Integer       ' signed
   yMin As Integer       ' signed
   xMax As Integer       ' signed
   yMax As Integer       ' signed
   MacStyle As Integer
   LowestRecPPEM As Integer
   FontDirectionHint As Integer ' signed
   IndexToLocFormat As Integer  ' signed
   GlyphDataFormat As Integer   ' signed
End Type

Private Type WINDOWSMETRICS
   ' We're not interested in anything but the fsSelection
   ' member here. Most other members have been
   ' set to have the right number of bits only.
   Version As Integer
   xAvgCharWidth As Integer
   usWeightClass As Integer
   usWidthClass As Integer
   fsType As Integer              'Embedding licensing
   ySubscriptXSize As Integer
   ySubscriptYSize As Integer
   ySubscriptXOffset As Integer
   ySubscriptYOffset As Integer
   ySuperscriptXSize As Integer
   ySuperscriptYSize As Integer
   ySuperscriptXOffset As Integer
   ySuperscriptYOffset As Integer
   yStrikeoutSize As Integer
   yStrikeoutPosition As Integer
   sFamilySize As Integer
   panose(0 To 9) As Byte
   ulUnicodeRange1 As Long
   ulUnicodeRange2 As Long
   ulUnicodeRange3 As Long
   ulUnicodeRange4 As Long
   achVendID(0 To 3) As Byte
   fsSelection As Integer
   usFirstCharIndex As Integer
   usLastCharIndex As Integer
   sTypoAscender As Integer
   sTypoDescender As Integer
   sTypoLineGap As Integer
   usWinAscent As Integer
   usWinDescent As Integer
   ulCodePageRange1 As Long
   ulCodePageRange2 As Long
End Type

' Bit values within WINDOWSMETRICS.fsSelection field
Private Const bitItalic As Integer = 0
Private Const bitUnderscore As Integer = 1
Private Const bitNegative As Integer = 2
Private Const bitOutlined As Integer = 3
Private Const bitStrikeout As Integer = 4
Private Const bitBold As Integer = 5
Private Const bitRegular As Integer = 6
   
' The Naming Table is organized as follows:
Private Type NAMINGTABLE ' (page 77: "name - Naming Table")
   FormatSelector As Integer               ' =0
   NumberOfNameRecords As Integer          ' What is says
   OffsetToStringStorage As Integer        ' (from start of NamingTable).
   'n(NumberOfNameRecords) As NameRecord   ' storage area for the NameRecords
   'Variable(x) As Byte                    ' storage area for the string data
End Type

' Each NameRecord looks like this:
Private Type NAMERECORD ' (page 77: "name - Naming Table")
   PlatformID As Integer
   PlatformSpecificEncodingID As Integer
   LanguageID As Integer 'un
   NameID As Integer ' NameIDs
   StringLength As Integer 'un
   StringOffset As Integer ' (from start of storage area)
End Type

Private Const PlatformMacintosh As Integer = 1
Private Const PlatformWindows As Integer = 3

'The following name ID's are defined, and they apply to all platforms:
Private Enum NameIDs ' (page 81: "Name ID's")
   ' Name strings are stored by platform, encoding and language (placing
   ' separate strings in for both Apple and MS platforms)
   gfiCopyrightNotice = 0 ' Copyright notice
   gfiFontFamilyName = 1  ' Font Family name
   gfiFontSubfamilyName = 2 ' For purposes of definition, this is assumed
                         ' to address style (italic, oblique) and weight
                         ' (light, bold, black, etc.) only. A font with
                         ' no particular differences in weight or style
                         ' (e.g. medium weight, not italic and fsSelection
                         ' bit 6 set) should have the string "Regular"
                         ' stored in this position.
   gfiUniqueFontIdentifier = 3 ' Some unique ID for this font.
   gfiFullFontName = 4   ' this should simply be a combination of strings 1
                         ' and 2. Exception: if string 2 is "Regular," then
                         ' use only string 1. This is the font name that
                         ' Windows will expose to users.
   gfiVersionString = 5  ' Version string. Supposed to be in n.nn format,
                         ' but may include whatever (e.g. a date)
   gfiPostscriptName = 6 ' Postscript name for the font.
   gfiTrademark = 7      ' Trademark notice/information for this font.
End Enum
' ****************************************************************

' Member Variables
Private m_Broadcast As Boolean
Private m_FileNameTTF As String
Private m_FaceName As String
Private m_LocaleID As Long
Private m_WasInstalled As Boolean
Private m_ValidFontFile As Boolean
Private m_NameRecs() As NAMERECORD
Private m_Metrics As WINDOWSMETRICS
Private m_Strings() As String

' *********************************************
'  Initialization / Termination
' *********************************************
Private Sub Class_Initialize()
   ' Initialize member vars
   m_Broadcast = False
   m_LocaleID = GetSystemDefaultLCID()
End Sub

Private Sub Class_Terminate()
   ' Kill resource before exit.
   Call RemoveFont
End Sub

' *********************************************
'  Public Properties
' *********************************************
Public Property Let Broadcast(ByVal NewVal As Boolean)
   m_Broadcast = NewVal
End Property

Public Property Get Broadcast() As Boolean
   Broadcast = m_Broadcast
End Property

Public Property Let FontFile(ByVal NewVal As String)
   ' Kill old resource before creating new.
   Call RemoveFont
   
   ' If font doesn't appear to already be
   ' installed, add resource tempoarily.
   If IsFile(NewVal) Then
      m_FileNameTTF = NewVal
      Call RefreshFontInfo(m_FileNameTTF)
      m_WasInstalled = GetFaceName()
      If Not m_WasInstalled Then
         Call AddFont(m_FileNameTTF)
         Call GetFaceName
      End If
   End If
End Property

Public Property Get FontFile() As String
   FontFile = m_FileNameTTF
End Property

' *********************************************
'  Public Properties: Read-only
' *********************************************
Public Property Get Bold() As Boolean
   Bold = BitValueI(m_Metrics.fsSelection, bitBold)
End Property

Public Property Get Copyright() As String
   Copyright = GetFontInfo(gfiCopyrightNotice)
End Property

Public Property Get FaceName() As String
   FaceName = m_FaceName
End Property

Public Property Get FamilyName() As String
   FamilyName = GetFontInfo(gfiFontFamilyName)
End Property

Public Property Get FullName() As String
   FullName = GetFontInfo(gfiFullFontName)
End Property

Public Property Get Installed() As Boolean
   Installed = m_WasInstalled
End Property

Public Property Get Italic() As Boolean
   Italic = BitValueI(m_Metrics.fsSelection, bitItalic)
End Property

Public Property Get StrikeOut() As Boolean
   StrikeOut = BitValueI(m_Metrics.fsSelection, bitStrikeout)
End Property

Public Property Get SubFamilyName() As String
   SubFamilyName = GetFontInfo(gfiFontSubfamilyName)
End Property

Public Property Get PostscriptName() As String
   PostscriptName = GetFontInfo(gfiPostscriptName)
End Property

Public Property Get Trademark() As String
   Trademark = GetFontInfo(gfiTrademark)
End Property

Public Property Get Underlined() As Boolean
   Underlined = BitValueI(m_Metrics.fsSelection, bitUnderscore)
End Property

Public Property Get UniqueIdentifier() As String
   UniqueIdentifier = GetFontInfo(gfiUniqueFontIdentifier)
End Property

Public Property Get ValidTrueTypeFile() As Boolean
   ValidTrueTypeFile = m_ValidFontFile
End Property

Public Property Get VersionString() As String
   VersionString = GetFontInfo(gfiVersionString)
End Property

' *********************************************
'  Public Methods
' *********************************************

' *********************************************
'  Private methods
' *********************************************
Private Function AddFont(ByVal FileNameTTF As String) As Boolean
   Dim nRet As Long
   ' Add resource to Windows font table.
   nRet = AddFontResource(FileNameTTF)
   If nRet Then
      ' Return success.
      AddFont = True
      ' Alert all running apps?
      If m_Broadcast Then
         Call SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, ByVal 0&)
      End If
   End If
End Function

Private Function GetFontInfo(Optional ByVal Info As NameIDs = gfiFullFontName) As String
   Dim i As Long
   Dim sRet As String
   ' Loop through name table records, looking for match.
   ' Order of preference:
   '   1st, Platform: Windows, Locale: System
   '   2nd, Platform: Windows
   '   3rd, Whatever
   For i = 0 To UBound(m_NameRecs)
      If m_NameRecs(i).NameID = Info Then
         ' If this is targetted for Windows, hang
         ' onto it.  Exit if also locale correct.
         If m_NameRecs(i).PlatformID = PlatformWindows Then
            sRet = m_Strings(i)
            If m_NameRecs(i).LanguageID = m_LocaleID Then
               Exit For
            End If
         Else
            ' If we don't have anything yet, hang onto
            ' this one in case nothing better comes.
            If Len(sRet) = 0 Then
               sRet = m_Strings(i)
            End If
         End If
      End If
   Next i
   GetFontInfo = sRet
End Function

Private Function RefreshFontInfo(ByVal FileNameTTF As String) As Boolean
   Dim TableOffsets As OFFSETTABLE
   Dim TableDir() As TABLEDIRECTORY
   Dim HeadTable As FONTHEADER
   Dim NameTable As NAMINGTABLE
   Dim lpTable As Long      ' Pointer to table
   Dim lpString As Long     ' Pointer to string data
   Dim hFile As Integer
   Dim i As Long
   Dim j As Long
   Const MagicNumber As Long = &H5F0F3CF5

   On Error GoTo Bail ' very basic error handler
   
   hFile = FreeFile()
   Open FileNameTTF For Binary Access Read Lock Write As #hFile
   
   ' Get the offset table
   Get #hFile, 1, TableOffsets
   TableOffsets.NumberOfTables = ToLittleEndianI(TableOffsets.NumberOfTables)
   
   ' Redim TableDir to hold all tables
   ReDim TableDir(0 To TableOffsets.NumberOfTables - 1)
   
   ' Get the table directory
   Get #hFile, , TableDir
   
   ' Check each table reference in the directory;
   ' use "head" table to confirm validity of file.
   i = 0
   Do
      If TableDir(i).Tag = "head" Then
         ' Found the right one, now get information from the table.
         lpTable = ToLittleEndianL(TableDir(i).Offset)
                                                            
         ' Read the table info. Note that VB uses 1-based file
         ' i/o, while the pointers we retrieve are 0-based.
         Get #hFile, lpTable + 1, HeadTable
         
         ' Check magic number, and leave loop.
         m_ValidFontFile = (MagicNumber = ToLittleEndianL(HeadTable.MagicNumber))
         Exit Do
      End If
   
      ' 0 To TableOffsets.NumberOfTables - 1
      i = i + 1
   Loop While i < TableOffsets.NumberOfTables
   
   If m_ValidFontFile Then
      ' Check each table reference in the directory;
      ' string info is stored in the "name" table.
      i = 0
      Do
         If TableDir(i).Tag = "name" Then
            ' Found the right one, now get information from the table.
            lpTable = ToLittleEndianL(TableDir(i).Offset)
                                                               
            ' Read the table header info, NameRecords
            ' follow immediately (VB is 1-based!).
            Get #hFile, lpTable + 1, NameTable
            NameTable.NumberOfNameRecords = ToLittleEndianI(NameTable.NumberOfNameRecords)
            NameTable.OffsetToStringStorage = ToLittleEndianI(NameTable.OffsetToStringStorage)
            ReDim m_NameRecs(0 To (NameTable.NumberOfNameRecords - 1))
            
            ' Read all the name records into array.
            For j = 0 To UBound(m_NameRecs)
               Get #hFile, , m_NameRecs(j)
               m_NameRecs(j).PlatformID = ToLittleEndianI(m_NameRecs(j).PlatformID)
               m_NameRecs(j).PlatformSpecificEncodingID = ToLittleEndianI(m_NameRecs(j).PlatformSpecificEncodingID)
               m_NameRecs(j).LanguageID = ToLittleEndianI(m_NameRecs(j).LanguageID)
               m_NameRecs(j).NameID = ToLittleEndianI(m_NameRecs(j).NameID)
               m_NameRecs(j).StringLength = ToLittleEndianI(m_NameRecs(j).StringLength)
               m_NameRecs(j).StringOffset = ToLittleEndianI(m_NameRecs(j).StringOffset)
            Next j
            
            ' Get associated string data for each name record.
            ReDim m_Strings(0 To UBound(m_NameRecs))
            For j = 0 To UBound(m_NameRecs)
               If m_NameRecs(j).StringLength Then
                  ' Calculate the string starting position: From the start of
                  ' the naming table, add the offset to the string storage area
                  ' and add the offset for this specific value.
                  lpString = lpTable + _
                     NameTable.OffsetToStringStorage + _
                     m_NameRecs(j).StringOffset
                  
                  ' Provide the buffer neccessary to hold the data.
                  m_Strings(j) = Space$(m_NameRecs(j).StringLength)
                  
                  ' Read and convert string if Unicode.
                  Get #hFile, lpString + 1, m_Strings(j)
                  If (Asc(m_Strings(j)) = 0) Then
                     ' The string is Unicode indicated by the
                     ' leading NULL, so shift it left one byte,
                     ' insert a NULL at its end, and convert.
                     m_Strings(j) = StrConv(Right$(m_Strings(j), m_NameRecs(j).StringLength - 1) & vbNullChar, vbFromUnicode)
                  End If
               End If
            Next j
         
         ' Also grab the Windows Metrics info while here.
         ElseIf TableDir(i).Tag = "OS/2" Then
            ' Found the right one, now get information from the table.
            ' Store the start address of this table for later use.
            lpTable = ToLittleEndianL(TableDir(i).Offset)
         
            ' Read the table info (VB is 1-based!).
            ' fsSelection contains info on Bold, Italics, etc.
            Get #hFile, lpTable + 1, m_Metrics
            m_Metrics.fsSelection = ToLittleEndianI(m_Metrics.fsSelection)
         End If
         
         ' 0 To TableOffsets.NumberOfTables - 1
         i = i + 1
      Loop While i < TableOffsets.NumberOfTables
      
      ' Return success
      RefreshFontInfo = True
   End If  ' m_ValidFontFile
      
Bail:
   Close #hFile
End Function

Private Function GetFaceName() As Boolean
   Dim fnt As New StdFont
   Dim FaceName(0 To 2) As String
   Dim i As Long
   
   ' FaceName may be either FamilyName, a
   ' combination of FamilyName and SubFamilyName,
   ' or FullName, so we need to try all.
   FaceName(0) = GetFontInfo(gfiFontFamilyName)
   FaceName(1) = GetFontInfo(gfiFontFamilyName) & " " & _
                 GetFontInfo(gfiFontSubfamilyName)
   FaceName(2) = GetFontInfo(gfiFullFontName)
   
   ' Attempt to select same named font into StdFont
   ' object. If successful, we don't need to
   ' temporarily install this file to see it.
   On Error Resume Next
   For i = 0 To 2
      fnt.Name = FaceName(i)
      If fnt.Name = FaceName(i) Then
         m_FaceName = FaceName(i)
         GetFaceName = True
         Exit Function
      End If
   Next i
End Function

Private Function RemoveFont() As Boolean
   ' If there are outstanding references to a
   ' font, the associated resource remains loaded
   ' until no device context is using it.
   If Len(m_FaceName) Then
      If Not m_WasInstalled Then
         Call RemoveFontResource(m_FileNameTTF)
         Debug.Print "Removed " & m_FaceName & "!!!"
         ' Alert all running apps?
         If m_Broadcast Then
            Call SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, ByVal 0&)
         End If
         ' Reset some vars.
         m_FaceName = ""
         m_ValidFontFile = False
      End If
   End If
End Function

' *********************************************
'  Private bit-twiddling methods
' *********************************************
Private Function ToLittleEndianI(ByRef Value As Integer) As Integer
   ToLittleEndianI = ByteSwap(Value)
End Function

Private Function ToLittleEndianL(ByRef Value As Long) As Long
   ToLittleEndianL = MakeLong(ByteSwap(WordLo(Value)), _
                              ByteSwap(WordHi(Value)))
End Function

Private Function BitValueI(ByVal WordIn As Integer, ByVal bit As Integer) As Boolean
   If bit >= 0 And bit <= 15 Then
      ' Return Nth power bit as true/false
      BitValueI = ((WordIn And (2 ^ bit)) > 0)
   Else
      ' Could raise an error, if more appropriate?
      BitValueI = False
   End If
End Function

Private Function ByteSwap(ByVal WordIn As Integer) As Integer
   Dim ByteHi As Integer
   Dim ByteLo As Integer
   Dim NewHi As Long
   
   ' Separate bytes using same strategy as in
   ' ByteHi and ByteLo functions. Faster to do
   ' it inline than to make function calls.
   If WordIn < 0 Then
      ByteHi = (WordIn + &H10000) \ &H100
   Else
      ByteHi = WordIn \ &H100
   End If
   ByteLo = WordIn And &HFF
   
   ' Shift low byte left by 8
   NewHi = ByteLo * &H100&
   
   ' Account for sign-bit
   If NewHi > &H7FFF Then
      ByteLo = NewHi - &H10000
   Else
      ByteLo = NewHi
   End If
   
   ' Place high byte in low position
   ByteSwap = ByteLo Or ByteHi
End Function

Private Function MakeLong(ByVal WordHi As Integer, ByVal WordLo As Integer) As Long
   ' High word is coerced to Long to allow it to
   ' overflow limits of multiplication which shifts
   ' it left.
   MakeLong = (CLng(WordHi) * &H10000) Or (WordLo And &HFFFF&)
End Function

Private Function WordHi(ByVal LongIn As Long) As Integer
   ' Mask off low word then do integer divide to
   ' shift right by 16.
   WordHi = (LongIn And &HFFFF0000) \ &H10000
End Function

Private Function WordLo(ByVal LongIn As Long) As Integer
   ' Low word retrieved by masking off high word.
   ' If low word is too large, twiddle sign bit.
   If (LongIn And &HFFFF&) > &H7FFF Then
      WordLo = (LongIn And &HFFFF&) - &H10000
   Else
      WordLo = LongIn And &HFFFF&
   End If
End Function

' *********************************************
'  Private file-related methods
' *********************************************
Private Function IsFile(SpecIn As String) As Boolean
   Dim Attr As Long
   
   ' Guard against bad SpecIn by ignoring errors.
   On Error Resume Next
   
   ' Get attribute of SpecIn.
   Attr = GetAttr(SpecIn)
   If Err = 0 Then
      ' No error, so something was found.
      ' If Directory attribute set, then not a file.
      If (Attr And vbDirectory) = vbDirectory Then
         IsFile = False
      Else
         IsFile = True
      End If
   End If
End Function


