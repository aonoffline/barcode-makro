Public Sub InsertBarcode


	Dim Range As Object
	Dim Flags As Long
	Dim stil As Variant
	Dim hefo As Variant
	
	myDoc = thisComponent
	mySheet = myDoc.sheets(0)
	mySheet2 = myDoc.sheets(1)
	
    mycell2 = mysheet.getCellByPosition(6,1)
	runner=mycell2.value
	
	Range = mySheet2.getCellRangeByName("A1:DA70")

	Flags = com.sun.star.sheet.CellFlags.VALUE + _
    	com.sun.star.sheet.CellFlags.DATETIME + _
	    com.sun.star.sheet.CellFlags.STRING + _
	    com.sun.star.sheet.CellFlags.OBJECTS + _
	    com.sun.star.sheet.CellFlags.EDITATTR

	Range.clearContents(Flags)
	MsgBox "Clean&GO"
	
	stil=myDoc.StyleFamilies.getByName("PageStyles")
	hefo=stil.getByName(mySheet2.PageStyle)
	hefo.HeaderOn=FALSE
	hefo.FooterOn=FALSE
	
	hefo.BottomMargin = 0
  	hefo.LeftMargin = 0
  	hefo.RightMargin = 0
  	hefo.TopMargin = 1000
	
	Dim args(8) as new com.sun.star.beans.NamedValue
	Dim yjumper as Long
	Dim xjumper as Long
	yjumper = 1000
	xjumper = 1100
	Dim initiatior As Long
	For initiatior = 1 To 3 Step + 1
	For m = 7 To 0 Step - 1

	
 	Dim positioner As Long
 	positioner = positioner + 1
 	mycell3 = mysheet.getCellByPosition(0 ,positioner)
 	jumper = mycell3.string
    mycell = mysheet.getCellByPosition(1, positioner)
	Barcodewert = mycell.string
	'MsgBox Barcodewert
	Dim oJob as Object
	oJob = createUnoService("org.libreoffice.Barcode")
	
    
    args(0).Name = "Action"
    args(0).Value = "InsertBarcode"
    args(1).Name = "BarcodeType"
    args(1).Value = "CODE128"
    args(2).Name = "BarcodeValue"
    args(2).Value = Barcodewert
    args(3).Name = "BarcodeAddChecksum"
    args(3).Value = True
    args(4).Name = "WidthScale"
    args(4).Value = "45"
    args(5).Name = "HeightScale"
    args(5).Value = "55"
    args(6).Name = "PositionX"
    args(6).Value = xjumper
    args(7).Name = "PositionY"
    args(7).Value = yjumper
    args(8).Name = "TargetComponent"
    args(8).Value = ThisComponent
    oJob.execute(args)
    yjumper = yjumper +3500	 



Next m
		xjumper = xjumper + 6900
		yjumper = 1000
Next initiatior

End  Sub

#zweckform 3490 70x36mm (DINA4)
