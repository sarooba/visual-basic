Private Sub Add_Click() Adodc1.Recordset.AddNew Textl.SetFocus End Sub 
DELETE: 
Private Sub Delete_Click() 
If MsgBox ("DELETE IT?",vbOKCancel)= vbOK Then 
Adodc1.Recordset.Delete End If MsgBox "ONE ROW DELETED" Textl.Text- " " 
Text2.Text - " " 
Text3.Text - " " 
Text4.Text - " " 
Text5.Text - " " 
Text6.Text - " " 
Text7.Text - " " 
Text8.Text - " " 
Text9.Text - " " Textl0.Text - " " End Sub SAVE: 
Private Sub Save_Click() 
If MsgBox ("SAVE IT?",vbOKCancel ) = vbOK Then Adodc1.Recordset.Update Else 
Adodc1.Recordset.CancelUpdate 
End If End Sub FIND: 
Private Sub Find_Click() Dim N as string 
N = InputBox ("Enter the accno") Adodc1.Recordset.Find "accno=" & N 
If Adodcl.Recordset.BOF or Adodc1.Recordset.EOF Then MsgBox "Record not found" 
End If End Sub 
Department of Computer Science and Engineering                                                    
Page:______    
K. Ramakrishnan College of Engineering (Autonomous), Trichy 
UPDATE: 
Private Sub Update_Click() Adodc1.Recordset.EditMode Adodc1.Recordset.Update End 
Sub FIRST: 
Private Sub First_Click() Adodc1.Recordset.MoveFirst End Sub LAST: 
Private Sub Last_Click() Adodc1.Recordset.MoveLast End Sub NEXT: 
Private Sub Next_Click() Adodc1.Recordset.MoveNext End Sub PREVIOUS: 
Private Sub Previous_Click() Adodc1.Recordset.MovePrevious End Sub DEPOSIT: 
Private Sub Deposit_Click0 Dim N1 as string 
N = InputBox ("Enter the accno") Adodcl.Recordset.Find "accno=" & N Nl = InputBox 
("Enter the amount") Text4.Text= val (Text4.Text) + Nl Adodc1.Recordset.Update 
End Sub 
WITHDRAW: 
Private Sub Withdraw_Click() Dim Nl as string 
N = InputBox ("Enter the accno") Adodcl.Recordset.Find "accno=" & N Nl = InputBox 
("Enter the amount") Text4.Text= val (Text4.Text) - NlAdodcl.Recordset.Update 
End Sub EXIT: 
Private Sub Add_Click() Unload Me End Sub FUNCTION: 
Function Calculate() 
Text8.Text=val(Text4.Text) + val (Text5.Text) + val (Text6.Text) + val (Text7.Text) 
Text9.Text=val(Text5.Text) + val (Text6.Text) + val (Text7.Text) 
Text 
10.Text=val(Text8.Text) 
+ 
val 
(Text9.Text) 
BASICPAY,HRA,DA,MA,GROSSPAY,DEDUCTION,NETPAY: 
Private Sub Basicpay_Change() Call Calculate End Sub 
Private Sub HRA_Change() Call Calculate End Sub 
Private Sub DA_Change() Call Calculate End Sub 
Private Sub MA_Change() Call Calculate 
End Sub 
Private Sub Grosspay_Change() Call Calculate End Sub 
End 
Function 
Department of Computer Science and Engineering                                                    
Page:______    
K. Ramakrishnan College of Engineering (Autonomous), Trichy 
Private Sub Deduction_Change() Call Calculate End Sub 
Private Sub Netpay_Change() Call Calculate End Sub 
