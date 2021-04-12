VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SectionHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Objects")
Option Explicit

Private m_ws As Worksheet
Private m_titles As Collection
Private m_sections As Collection

Property Get sections() As Collection
    Set sections = m_sections
End Property

Public Sub SetProperties(ByVal ws As Worksheet)
    Set m_ws = ws
    Set m_titles = GetSectionTitles
    Set m_sections = GetSections
End Sub

Private Function GetSectionStartRow(ByVal sectionTitle As String) As Range
    Dim rowNumber As Long
    Dim lookupTable As Range
    Set lookupTable = Range("B1:B5000")
    rowNumber = Application.WorksheetFunction.Match(sectionTitle, lookupTable, 0)
    Set GetSectionStartRow = Range("B" & rowNumber)
End Function

Private Function GetWholeSectionRange(ByVal sectionStart As Range) As Range
    Dim sectionStartRow As Long
    Dim sectionLastRow As Long
    
    sectionStartRow = sectionStart.Row
    sectionLastRow = sectionStart.End(xlDown).Row
    Dim result As Range
    Set result = Range("A" & sectionStartRow, "M" & sectionLastRow)
    Set GetWholeSectionRange = result
End Function

Private Function GetSectionTitles() As Collection
    Dim titles As Collection
    Set titles = New Collection
    titles.Add "1_Sekcja 1"
    titles.Add "1_Sekcja 2"
    titles.Add "1_Sekcja 3"
    Set GetSectionTitles = titles
End Function

Private Function GetSections() As Collection
    Dim sections As Collection
    Set sections = New Collection
    Dim title As Variant
    For Each title In m_titles
        Dim sectionStart As Range
        Set sectionStart = GetSectionStartRow(CStr(title))
        
        Dim wholeSectionRange As Range
        Set wholeSectionRange = GetWholeSectionRange(sectionStart)
        
        Dim sec As Section
        Set sec = Create.Section(wholeSectionRange)
        
        sections.Add sec
        Debug.Print sec.ToString
    Next
    
    Set GetSections = sections
End Function
