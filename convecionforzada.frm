VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} convecionforzada 
   Caption         =   "conveccion forzada"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8565.001
   OleObjectBlob   =   "convecionforzada.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "convecionforzada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function hfun(Pr, Re, k, l)
hfun = 0.332 * (Pr ^ (1 / 3)) * (Re ^ (1 / 2)) * k / l
End Function

Private Sub aguabonton_Click()
        Me.viscosidad.Value = 0.000979
        Me.densidad.Value = 998
        Me.capacidadc.Value = 4180
        Me.conductividad.Value = 0.601
End Sub

Private Sub CommandButton1_Click()
Velocidad.Text = ""
longitud.Text = ""
viscosidad.Text = ""
densidad.Text = ""
capacidadc.Text = ""
temperaturaf.Text = ""
temperaturap.Text = ""
conductividad.Text = ""
longitud.SetFocus
End Sub

Private Sub CommandButton2_Click()
    'se aclaran las variables
'Dim u, l, k, Ts, Tp, Cp, v, p, Pr, Re, dT, kl, Nu, h As Integer

    'se definen las ecuaciones para Pr y Re y kl


u = Me.viscosidad.Value
v = Me.Velocidad.Value
l = Me.longitud.Value
p = Me.densidad.Value
Cp = Me.capacidadc.Value
Tf = Me.temperaturaf.Value
Tp = Me.temperaturap.Value
k = Me.conductividad.Value
    
    
    
    'If aguaboton = True Then
        'Me.viscosidad.Value = 0.000979
        'Me.densidad.Value = 998
        'Me.capacidadc.Value = 4180
        'Me.conductividad.Value = 0.601
   ' End If
    
'se agrega un aviso cuando uno de los denominadores de una ecuación es 0


If Me.viscosidad.Value = "" Then
        MsgBox "Division por 0 no valida", vbExclamation
        End
End If

If Me.conductividad.Value = "" Then
        MsgBox "Division por 0 no valida", vbExclamation
        End
End If

If Me.longitud.Value = "" Then
        MsgBox "Division por 0 no valida", vbExclamation
        End
End If


Re = l * v * p / u

MsgBox Re


Pr = u * Cp / k
MsgBox Pr


'kl = k / l
h = hfun(Pr, Re, k, l)


'Nu = h * l / k
Me.hresul.Value = h

dT = Tp - Tf
MsgBox dT


q = h * dT
Me.Qecuacion.Value = q
    
End Sub

Private Sub CommandButton3_Click()
End
End Sub




Private Sub UserForm_Click()
Velocidad.Text = ""
longitud.Text = ""
viscosidad.Text = ""
densidad.Text = ""
capacidadc.Text = ""
temperaturaf.Text = ""
temperaturap.Text = ""
conductividad.Text = ""
longitud.SetFocus
End Sub
