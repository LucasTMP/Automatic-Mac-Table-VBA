Private Sub BT_Login_Click()

If pwd <> "" Then

WB1.Document.GetElementById("UserName").innerText = "USER"
WB1.Document.GetElementById("Password").innerText = pwd.Value

With WB1.Document
    Set elems = .getElementsByTagName("input")
    For Each e In elems
        If (e.getAttribute("value") = "Apply") Then
            e.Click
            Exit For
        End If
    Next e
End With
    pwd.Value = ""
    
    Do While WB1.ReadyState <> READYSTATE_COMPLETE
      DoEvents
      Loop
    
    WB1.Navigate "http://192.168.0.1/?lan_dhcp"
    
    Do While WB1.ReadyState <> READYSTATE_COMPLETE
      DoEvents
      Loop
    
    If WB1.LocationURL = "http://192.168.0.1/?lan_dhcp" Then
    BT_PegarTabela.Visible = True
    BT_Login.Visible = False
    pwd.PasswordChar = ""
    pwd.Value = "Conexão ativa"
    pwd.Locked = True
    Label1.Caption = "Status da Conexão :"
    End If
 
 
Else

MsgBox "Preencha o campo da senha.", vbInformation, "Senha não preenchida"


End If

End Sub



Private Sub BT_PegarTabela_Click()

Web_Table_Option_Two

Logoff

BT_Sair.Visible = True

pwd.Value = "Conexão fechada"
pwd.Locked = True
Label1.Caption = "Status da Conexão :"

End Sub


Private Sub BT_Sair_Click()

Unload Me

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub UserForm_Initialize()
WB1.Navigate "http://192.168.0.1"
pwd.Value = ""
pwd.PasswordChar = "*"
BT_PegarTabela.Visible = False
BT_Login.Visible = True
WB1.Visible = True
pwd.Locked = False
BT_Sair.Visible = False
End Sub


Function Web_Table_Option_Two()

 Dim htm As Object
    Dim Tr As Object
    Dim Td As Object
    Dim Tab1 As Object
    Dim retira3 As String
    Dim retira2 As String
    Dim retira1 As String
    Dim compara_txt As String
    
    Column_Num_To_Start = 5
    iRow = 4
    iCol = Column_Num_To_Start
    iTable = 4


Sheets("Controle Rede Local").Select

    'Loop Through Each Table and Download it to Excel in Proper Format
    'For Each Tab1 In WB1.Document.getElementsByTagName("table")
        With WB1.Document.getElementsByTagName("table")(iTable)
            For Each Tr In .Rows
            
            If InStr(Tr.Cells(0).innerText, ":") = 0 Then
            
            'retira1 = Tr.Cells(0).innertext
            retira2 = Tr.Cells(3).innerText
            retira3 = Tr.Cells(4).innerText
            
                For Each Td In Tr.Cells
                compara_txt = Td.innerText
                If StrComp(retira1, compara_txt) <> 0 Or Null Then
                    If StrComp(retira2, compara_txt) <> 0 Or Null Then
                      If StrComp(retira3, compara_txt) <> 0 Or Null Then
                    Sheets("Controle Rede Local").Cells(iRow, iCol).Select
                    Sheets("Controle Rede Local").Cells(iRow, iCol) = Td.innerText
                    iCol = iCol + 1
                    End If
                    End If
                    End If
                Next Td
                End If
            
                iCol = Column_Num_To_Start
                iRow = iRow + 1
            Next Tr
        End With

Table_visual
        'iTable = iTable + 1
        'iCol = Column_Num_To_Start
        'iRow = iRow + 1
    'Next Tab1

    MsgBox "Tabela DHCP Atualizada.", vbOKOnly, "Operação concluida"
    
End Function


Function Table_visual()


Sheets("Controle Rede Local").Range("tabela_dhcp").Select


   With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0
        .Weight = xlThin
        
    End With

End Function

Function Logoff()

With WB1.Document

    Set elems = .getElementsByTagName("a")
    For Each e In elems

        If (e.getAttribute("href") = "http://192.168.0.1/router.html") Then
            elems.Item(2).Click
            Exit For
            End If
            
    Next e

End With
End Function
