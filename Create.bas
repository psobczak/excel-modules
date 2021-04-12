Attribute VB_Name = "Create"
'@Folder("Object factory")
Option Explicit

Public Function Section(ByVal sectionRange As Range) As Section
    Dim sec As Section
    Set sec = New Section
    sec.SetProperties sectionRange:=sectionRange
    Set Section = sec
End Function

Public Function SectionHandler(ByVal ws As Worksheet) As SectionHandler
    Dim handler As SectionHandler
    Set handler = New SectionHandler
    handler.SetProperties ws:=ws
    Set SectionHandler = handler
End Function


