Sub Main
	Call GapDetection()	'Ejemplo-Detalle de ventas.IMD
End Sub


' An�lisis: Detectar omisiones de car�cter
Function GapDetection
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.Gaps
	task.FieldToUse =  "NUM_FACT"
	task.Mask = "NNNNNNN"
	dbName = "Omisiones_01.IMD"
	task.OutputDBName = dbName
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function