//Obtem o nome do computador

//obtem objeto que representa o servico WMI da máquina
var objWMIService = GetObject( "winmgmts://./root/cimv2" )

//obtem a lista de sistemas windows da máquina
var colItems = objWMIService.ExecQuery( "Select * from Win32_ComputerSystem", null , 48 )

//obtem o iterador da coleçao
var colProps = new Enumerator(colItems);

//cria variavel que receberá nome do computador
var pcName;

//itera sobre a coleção de dados retornada do servico WMI
for ( ; !colProps.atEnd(); colProps.moveNext()) { 
	//obtem um item da coleção
	p = colProps.item();
	//extrai o nome do computador do item
	pcName = p.name
}

//exibe o nome do computador
WScript.Echo ("nome do computador:"+pcName);
