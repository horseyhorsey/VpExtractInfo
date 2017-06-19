# VpExtractInfo
Visual basic script to extract Visual pinball collections to JSon, yaml or LEd config.

##How to use
Place the VpExtractInfo.vbs in your Visual Pinball scripts folder.

Exporting collections make use of the TimerInterval on Vp objects to specify the number for the sw, lamp, coil, etc.
If you have a lamp called "shootAgain" you would add the number of this lamp into the TimerInterval box.

To be able to call the functions add the following to a VP script.

	Option Explicit
	Randomize

	On Error Resume Next
	ExecuteGlobal GetTextFile("VpExtractInfo.vbs")
	If Err Then MsgBox "Can't open VpExtractInfo.vbs"
	On Error Goto 0

Export to yaml for p-roc passing in your table name, PR type, and VP collection name.

    PrintCollectionYaml "MyTableName", "PRLamps", Lamps
    PrintCollectionYaml "MyTableName", "PRSwitches", Switches
    PrintCollectionYaml "MyTableName", "PRCoils", Coils
    
    
Export to a basic config for use in LedShowEditor.
    
    PrintCollectionLedShowJSON "MyTableName", Lamps
    

Export Full list and safe.

    PrintCollectionFull "MyTableName", "PRLamps", Lamps
    PrintCollectionSafe "MyTableName", "PRLamps", Lamps


