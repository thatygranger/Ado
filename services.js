var objWMIService = GetObject( "winmgmts://./root/cimv2" )

var colItems = objWMIService.ExecQuery("Select * from Win32_Service	",null,48)
var colProps = new Enumerator(colItems);
var serviceArray = new Array ();

for ( ; !colProps.atEnd(); colProps.moveNext()) { 
	p = colProps.item();
	var obj = new Object ()
	serviceArray.push (p);
}


for (var i = 0; i< serviceArray.length; i+=1) {
	var service = serviceArray[i];
	WScript.Echo ("nome: "+service.Name );
	WScript.Echo ("nome fantasia: "+service.DisplayName );
	WScript.Echo ("status: "+service.State  );
	WScript.Echo  ();
}


