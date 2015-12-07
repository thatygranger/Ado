/*
- Criar uma função que retorna a quantidade de Memória RAM da máquina

Fora da função, se o valor da memória for menor que 2gb e a máquina for do tipo laptop, 
adicionar no relatório que o laptop está fora do perfil de hardware da empresa. 
Caso seja maior que 2gb, imprimir no relatório que a máquina está aderente ao perfil de hardware da empresa.

Fora da função, se o valor da memória for menor que 4gb e a máquina for do tipo desktop, 
adicionar no relatório que o desktop está fora do perfil de hardware da empresa. 
Caso seja maior que 4gb, imprimir no relatório que a máquina está aderente ao perfil de hardware da empresa.
*/

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
