VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConvecciónNatural 
   Caption         =   "UserForm1"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7770
   OleObjectBlob   =   "ConvecciónNatural.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ConvecciónNatural"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub calcular_Click()
'Ingresar datos
Longitudplaca = Me.LONGITUD.Value
anchoplaca = Me.anchoplaca.Value
TemperaturaPlaca = Me.temperaturap.Value
Temperaturaaire = Me.temperaturaa.Value
conductividad = Me.conductividad.Value
viscosidad = Me.viscosidad.Value
NumeroPrandtl = Me.prandtl.Value

'Calculos
Tf = Temperaturaaire + TemperaturaPlaca / 2

B = 1 / Tf

NúmeroGrashof = 9.8 * B * Temperaturaaire - TemperaturaPlaca * Longitudplaca ^ 3 / viscosidad ^ 2

Ra = NumeroGrashof * NumeroPrandtl

'Advertencia para el valor de Nu dependiendo el valor de Ra

If Ra < 10 ^ 11 Then Nu = 0.68 + (0.67 * Ra ^ 1 / 4) / (1 + (0.492 / NumeroPrandtl) ^ 9 / 16) ^ 4 / 9

If 10 ^ 11 < Ra < 10 ^ 12 Then Nu = 0.825 + (0.387 * Ra ^ 1 / 6) / (1 + (0.492 / NumeroPrandtl) ^ 9 / 16) ^ 8 / 27

'Hallamos h

h = Nu * conductividad / Longitudplaca

'Hallamos q

q = h * (Temperaturaaire - TemperaturaPlaca)

'Salida de datos

Me.hnumero.Value = h
Me.qnumero.Value = q

End Sub

Private Sub CommandButton1_Click()
End
End Sub

Private Sub UserForm_Click()
Me.LONGITUD.Value = ""
Me.anchoplaca.Value = ""
Me.temperaturap.Value = ""
Me.temperaturaa.Value = ""
Me.conductividad.Value = ""
Me.viscosidad.Value = ""
Me.prandtl.Value = ""
End Sub
