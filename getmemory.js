var objWMIService = GetObject( "winmgmts://./root/cimv2" )
var colItems = objWMIService.ExecQuery( "Select * from Win32_ComputerSystem", null , 48 )
var colProps = new Enumerator(colItems);
var pcName;
for ( ; !colProps.atEnd(); colProps.moveNext()) { 
	p = colProps.item();
	pcName = p.name
}

var colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",null,48)
var colProps = new Enumerator(colItems);
var totalMemory = 0;;
for ( ; !colProps.atEnd(); colProps.moveNext()) { 
	p = colProps.item();
	totalMemory += ( p.Capacity/1048576 );
}
WScript.Echo ("memoria total: "+totalMemory+" mb");
