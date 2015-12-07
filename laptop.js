var myComputer = ".";
var objWMIService = GetObject( "winmgmts:\\\\" + myComputer + "\\root\\cimv2" );
var colItems = objWMIService.ExecQuery( "Select * from Win32_Battery" );
var IsLaptop = false;
var objItem = new Enumerator(colItems);
for (;!objItem.atEnd();objItem.moveNext()) {
	IsLaptop = true;
}
	
if (IsLaptop)
	WScript.Echo ("eh laptop");
else
	WScript.Echo ("eh desktop");

	
