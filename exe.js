//função 1 - Nome do computador

var nomepc = function (){

	//obtem objeto que representa o servico WMI da máquina
	var objWMIService = GetObject( "winmgmts://./root/cimv2" );
	
	//obtem a lista de sistemas windows da máquina
	var colItems = objWMIService.ExecQuery( "Select * from Win32_ComputerSystem", null , 48 );
	
	//obtem o iterador da coleçao
	var colProps = new Enumerator(colItems);
	
	//cria variavel que receberá nome do computador
	var pcName;
	
		//itera sobre a coleção de dados retornada do servico WMI
		for ( ; !colProps.atEnd(); colProps.moveNext()) { 
			//obtem um item da coleção
			p = colProps.item();
			//extrai o nome do computador do item
			pcName = p.name;
		}
        return pcName;
}

//--------------------------------------------------------------------------//

//Função 2 - Desktop ou Laptop


var tipoeq = function (){

	var myComputer = ".";
	
	var objWMIService = GetObject( "winmgmts:\\\\" + myComputer + "\\root\\cimv2" );
	
	var colItems = objWMIService.ExecQuery( "Select * from Win32_Battery" );
	
	var IsLaptop = false;
	
	var objItem = new Enumerator(colItems);
	
		for (;!objItem.atEnd();objItem.moveNext()) {
			IsLaptop = true;
		}
		/*	
		if (IsLaptop)
			WScript.Echo ("O equipamento eh laptop");
		else
			WScript.Echo ("O equipamento eh desktop");
	         */		
	return IsLaptop;
}

//--------------------------------------------------------------------------
//função 3 - Quantidade de memória


var qtdmemoria = function (){

	var objWMIService = GetObject( "winmgmts://./root/cimv2" );
	var colItems = objWMIService.ExecQuery( "Select * from Win32_ComputerSystem", null , 48 );
	var colProps = new Enumerator(colItems);
	var pcName;
	
		for ( ; !colProps.atEnd(); colProps.moveNext()) { 
			p = colProps.item();
			pcName = p.name;
		}
	
	var colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",null,48);
	var colProps = new Enumerator(colItems);
	var totalMemory = 0;
	
		for ( ; !colProps.atEnd(); colProps.moveNext()) { 
			p = colProps.item();
			totalMemory += ( p.Capacity/1048576 );
		}
	WScript.Echo ("memoria total: "+totalMemory+" mb");
return totalMemory;
}

//--------------------------------------------------------------------------

//Função 5 Serviços rodando na máquina

var servicos = function (){
var objWMIService = GetObject( "winmgmts://./root/cimv2" );

var colItems = objWMIService.ExecQuery("Select * from Win32_Service	",null,48);
var colProps = new Enumerator(colItems);
var serviceArray = new Array ();

for ( ; !colProps.atEnd(); colProps.moveNext()) { 
	p = colProps.item();
	var obj = new Object ();
	serviceArray.push (p);
}


for (var i = 0; i< serviceArray.length; i+=1) {
	var service = serviceArray[i];
	WScript.Echo ("nome: "+service.Name );
	WScript.Echo ("nome fantasia: "+service.DisplayName );
	WScript.Echo ("status: "+service.State  );
	WScript.Echo  ();
}
return serviceArray;
}

//--------------------------------------------------------------------------
//Função 7 serviços iniciados com o S.O

var processo = function (){
var objWMIService = GetObject( "winmgmts://./root/cimv2" );

var colItems = objWMIService.ExecQuery("Select * from Win32_StartupCommand	",null,48);
var colProps = new Enumerator(colItems);
var processArray = new Array ();

for ( ; !colProps.atEnd(); colProps.moveNext()) { 
	p = colProps.item();
	var obj = new Object ();
	processArray.push (p);
}


for (var i = 0; i< processArray.length; i+=1) {
	var process = processArray[i];
	WScript.Echo ("nome: "+process.Name );
}
return processArray;
}


//--------------------------------------------------------------------------\\



WScript.Echo("\n");
WScript.Echo("Relatório de Manutenção");
WScript.Echo("\n");


//Função 1
WScript.Echo("Nome do computador: " + nomepc());


//Função 2
if (tipoeq())
	WScript.Echo ("O equipamento eh laptop");
	else	
	WScript.Echo ("O equipamento eh desktop");


//Função 3
if(tipoeq()){
	if (qtdmemoria() < 2048)
	WScript.Echo("Equipamento fora das especificações da empresa");
	else
	WScript.Echo("Equipamento de acordo com as especificações da empresa");
	
}
else {
if (qtdmemoria() < 4096)
	WScript.Echo("Equipamento fora das especificações da empresa");
	else
	WScript.Echo("Equipamento de acordo com as especificações da empresa");
}

//Função 5 - WinDefend, sppsvc, MpsSvc
var varservice = servicos();
for (var i = 0; i< varservice.length; i+=1) {
		var varservice2 = varservice[i];
	if(varservice2.Name=="WinDefend"){
		if (varservice2.State=="Stopped")
		WScript.Echo("ALERTA DE SEGURANÇA! O serviço WINDEFEND não está rodando.");
		}
		if(varservice2.Name=="sppsvc"){
		if (varservice2.State=="Stopped")
		WScript.Echo("ALERTA DE SEGURANÇA! O serviço sppsvc não está rodando.");
		}
		if(varservice2.Name=="MpsSvc"){
		if (varservice2.State=="Stopped")
		WScript.Echo("ALERTA DE SEGURANÇA! O serviço MpsSvc não está rodando.");
		}
}

//Função 7 Verificar OneDrive

var varprocess = processo();


var i = 0;
do {
var varprocess2 = varprocess[i];

if(varprocess2.Name=="OneDrive")
		WScript.Echo("O OneDrive está na lista de inicialização");
		else
		WScript.Echo("O OneDrive não está na lista de inicialização");
i++;
} while (varprocess2.Name != "OneDrive");



/*
var varprocess = processo();
for (var i = 0; i< varprocess.length; i+=1) {
	var varprocess2 = varprocess[i];
	if(varprocess2.Name=="OneDrive")
		WScript.Echo("O OneDrive está na lista de inicialização");
		else
		WScript.Echo("O OneDrive não está na lista de inicialização");
}
*/
