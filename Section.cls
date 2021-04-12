VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Section"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Objects")
Option Explicit

Private m_wholeRange As Range
Private m_name As String
Private m_lastColumn As Range
Private m_firstColumn As Range
Private m_dataRange As Range

Public Property Get LastColumn() As Range
    Set LastColumn = m_lastColumn
End Property

Public Property Get FirstColumn() As Range
    Set FirstColumn = m_firstColumn
End Property

Public Property Get DataRange() As Range
    Set DataRange = m_dataRange
End Property

Public Property Get Name() As String
    Name = m_name
End Property

Public Sub SetProperties(ByVal sectionRange As Range)
    Set m_wholeRange = sectionRange
    m_name = Range("B" & m_wholeRange.Row).Value
    Set m_dataRange = Range("A" & m_wholeRange.Row + 2, "M" & m_wholeRange.Rows.Count + 1)
    Set m_lastColumn = m_dataRange.Columns.Item(1)
    Set m_firstColumn = m_dataRange.Columns.Item(m_wholeRange.Columns.Count)
End Sub

Public Function ToString() As String
    ToString = "Section(name: " & m_name & ", wholeRange: " & m_wholeRange.Address & ", dataRange: " & m_dataRange.Address & ")"
End Function



