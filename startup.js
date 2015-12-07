var objWMIService = GetObject( "winmgmts://./root/cimv2" )

var colItems = objWMIService.ExecQuery("Select * from Win32_StartupCommand	",null,48)
var colProps = new Enumerator(colItems);
var processArray = new Array ();

for ( ; !colProps.atEnd(); colProps.moveNext()) { 
	p = colProps.item();
	var obj = new Object ()
	processArray.push (p);
}


for (var i = 0; i< processArray.length; i+=1) {
	var process = processArray[i];
	WScript.Echo ("nome: "+process.Name );
}
