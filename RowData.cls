VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RowData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private basicDate_ As Date
Private fundCode_ As String
Private bftAccount_ As String
Private positionCurrency_ As String
Private breakMGM_ As Currency
Private status_ As String

Property Let BasicDate(basic As Date)
    basicDate_ = basic
End Property

Property Get BasicDate() As Date
    BasicDate = basicDate_
End Property

Property Let FundCode(code As String)
    fundCode_ = code
End Property

Property Get FundCode() As String
    FundCode = fundCode_
End Property

Property Let BFTAccount(account As String)
    bftAccount_ = account
End Property

Property Get BFTAccount() As String
    BFTAccount = bftAccount_
End Property

Property Let PositionCurrency(pos As String)
    positionCurrency_ = pos
End Property

Property Get PositionCurrency() As String
    PositionCurrency = positionCurrency_
End Property

Property Let BreakMGM(break As Currency)
    breakMGM_ = break
End Property

Property Get BreakMGM() As Currency
    BreakMGM = breakMGM_
End Property

Property Let Status(stat As String)
    status_ = stat
End Property

Property Get Status() As String
    Status = status_
End Property
