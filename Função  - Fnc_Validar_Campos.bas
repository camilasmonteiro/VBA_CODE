'-- La función siguiente fue desarrollada para realizar la validación de los campos obligatorios de formularios y sub-formularios creados en access (microsoft office), 
'-- evitando la repetición de código y mejorando el tiempo de ejecución del sistema. La función es una categoría simple, se puede adaptar a otros proyectos.
'-- Desarrollado por: Camila Monteiro. El editor de código utilizado: Visual Studio Code.

'-- The function below was developed to perform the validation of mandatory form fields and subforms created in access (microsoft office),
'-- avoiding code repetition and improving system runtime.The function is simple category, can be adapted for other projects. Developed by: Camila Monteiro.
'-- The code editor used: Visual Studio Code.

'-- Create a public function and save in a new module.
Public Function Fnc_Fields_Validate (Frm As Form) As Boolean

Fnc_Fields_Validate = True
    For Each Control In Frm.Controls
        If Control.Tag = "*" Then
            If Control.Value = "" Or IsNull (Control.Value) Then
                MsgBox "Please fill the required & Controle.Value & fild to save the record!",vbInformation, "Required Fields"
                Fnc_Fields_Validate = False
                Exit Function
            End If
        End If
    Next
End Function 

'-- Open the action of your field or form and call the public function. Show the example below:
Private Sub Form_AfterUpdate()

If Fnc_Fields_Validate(Form) = True Then
    Me.txt_name.Setfocus
Else
    MsgBox "New registered user sucessfully registered!",vbInformation,"New Register"
End If

End Sub 
