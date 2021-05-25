Attribute VB_Name = "Module1"
Sub PasaBinario(Num, Largo, StrBin)

a = Num
StrBin = ""
Do
    a1 = Int(a / 2)
    StrBin = Trim(Str(a - a1 * 2)) & StrBin
    a = a1
Loop Until a < 2
StrBin = Trim(Str(a)) & StrBin

For r = 1 To Largo
    StrBin = "0" & StrBin
Next r
StrBin = Right(StrBin, Largo)


End Sub


Sub LimpiaTabla()
On Error GoTo 1

frmKVDiagram.Command4.Visible = False
For r = 1 To 100
    Unload frmKVDiagram.Entrada(r)
    Unload frmKVDiagram.Salida(r)
Next r

1 End Sub
Sub descargaCamaras()

frmKVDiagram.Camara(r).Visible = False
frmKVDiagram.C1(r).Visible = False
frmKVDiagram.C2(r).Visible = False
frmKVDiagram.C3(r).Visible = False
frmKVDiagram.C4(r).Visible = False
frmKVDiagram.T1(r).Visible = False
frmKVDiagram.T2(r).Visible = False
frmKVDiagram.T3(r).Visible = False
frmKVDiagram.T4(r).Visible = False
frmKVDiagram.NOT(r).Visible = False
frmKVDiagram.TNOT(r).Visible = False

On Error GoTo 1
For r = 1 To 100
    Unload frmKVDiagram.Camara(r)
    Unload frmKVDiagram.C1(r)
    Unload frmKVDiagram.C2(r)
    Unload frmKVDiagram.C3(r)
    Unload frmKVDiagram.C4(r)
    Unload frmKVDiagram.T1(r)
    Unload frmKVDiagram.T2(r)
    Unload frmKVDiagram.T3(r)
    Unload frmKVDiagram.T4(r)
    Unload frmKVDiagram.NOT(r)
    Unload frmKVDiagram.TNOT(r)
    Unload frmKVDiagram.Conector(r)
Next r

frmKVDiagram.Label3.Visible = False
frmKVDiagram.output.Visible = False
frmKVDiagram.fondo.Visible = False

1 End Sub

