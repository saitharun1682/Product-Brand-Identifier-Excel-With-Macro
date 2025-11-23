Private Sub commandbutton_Click()
     Dim ws As Worksheet
     Dim nextRow As Long
     Dim ctrl As Control
     Set ws = ThisWorkbook.Sheets("macro")
     
     
     'Find the next empty row
     nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
     
     
     'save TextBox and ComboBox Value
      ws.Cells(nextRow, 1).Value = txtname.Value
      ws.Cells(nextRow, 2).Value = cmbage.Value
      ws.Cells(nextRow, 4).Value = cmbbrand.Value
      
      
    'save Gender from Optionbuttons
    If optmale.Value Then ws.Cells(nextRow, 3).Value = optmale.Caption
    If optfemale.Value Then ws.Cells(nextRow, 3).Value = optfemale.Caption
    If optother.Value Then ws.Cells(nextRow, 3).Value = optother.Caption
    
    'Save Checkbokes As Yes/No
    ws.Cells(nextRow, 5).Value = IIf(chkajio.Value, "Yes", "No")
    ws.Cells(nextRow, 6).Value = IIf(chkmyntra.Value, "Yes", "No")
    ws.Cells(nextRow, 7).Value = IIf(Chkamazon.Value, "Yes", "No")
    
    
    'clear from fields
    For Each ctrl In Me.Controls
        Select Case TypeName(ctrl)
            Case "TextBox", "ComboBox"
                ctrl.Value = " "
            Case "optionButton", "checkBox"
                ctrl.Value = False
        End Select
    Next ctrl

                                         
End Sub

Private Sub UserForm_Initialize()

'List ages 18-60 in the ComboBox

Dim i As Long


For i = 18 To 60
   cmbage.AddItem i
Next i

  With cmbbrand
     .AddItem "Adidas"
.AddItem "Allen Solly"
.AddItem "Anita Dongre"
.AddItem "Aurelia"
.AddItem "BEYOUNG"
.AddItem "Biba"
.AddItem "Cottonworld"
.AddItem "FabIndia"
.AddItem "Flying Machine"
.AddItem "Global Desi"
.AddItem "H&M"
.AddItem "Jockey"
.AddItem "John Players"
.AddItem "Kappa"
.AddItem "Lee"
.AddItem "Levi’s"
.AddItem "Louis Philippe"
.AddItem "Max Fashion"
.AddItem "Monte Carlo"
.AddItem "Mufti"
.AddItem "Nike"
.AddItem "Oxemberg"
.AddItem "Pantaloons"
.AddItem "Park Avenue"
.AddItem "Pepe Jeans"
.AddItem "Peter England"
.AddItem "Puma"
.AddItem "Raymond"
.AddItem "Reebok"
.AddItem "Roadster"
.AddItem "Tommy Hilfiger"
.AddItem "UCB (United Colors of Benetton)"
.AddItem "Van Heusen"
.AddItem "W for Woman"
.AddItem "Zara"

  End With
  
     

End Sub
