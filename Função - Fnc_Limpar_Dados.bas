'The function below was developed to clear forms fields (textbox and combobox) avoiding code repetition.

'La función siguinte es para limpiar los campos de los formularios (textbox y combobox) evitando la repeteción de códigos.

'-- Create a public function and save in a module.
Public Function Fnc_Clear_Fields (Frm As Form)

Dim CtrL    As Control

On Error Resume Next
For Each CtrL In Frm.Controls
    If TypeOf CtrL Is Textbox Or TypeOf CtrL Is Combobox Then
        CtrL = Empty
    End If
Next

Err.clear

End Function

'-- Open the action of your control form and call the public function. Show the example below:
Private Sub btn_name_Click()

Call Fnc_Clear_Fields(Form_Name)

End Sub