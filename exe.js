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
	
	//obtem objeto que representa o servico WMI da máquina	
	var objWMIService = GetObject( "winmgmts:\\\\" + myComputer + "\\root\\cimv2" );
	//obtem informações da bateria da maquina
	var colItems = objWMIService.ExecQuery( "Select * from Win32_Battery" );
	var IsLaptop = false;
	var objItem = new Enumerator(colItems);
	
	//verifica se existe alguma informação sobre a bateria do equipamento, se sim, IsLaptop = true
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

	//obtem objeto que representa o servico WMI da máquina	
	var objWMIService = GetObject( "winmgmts://./root/cimv2" );
	//obtem informações do sistema da maquina
	var colItems = objWMIService.ExecQuery( "Select * from Win32_ComputerSystem", null , 48 );
	var colProps = new Enumerator(colItems);
	var pcName;
	//busca na lista de informações do sistema o nome da maquina e armazena na variavel 'pcName'
		for ( ; !colProps.atEnd(); colProps.moveNext()) { 
			p = colProps.item();
			pcName = p.name;
		}
	//obtem informações sobre a memória física da maquina
	var colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",null,48);
	var colProps = new Enumerator(colItems);
	var totalMemory = 0;
		//verifica a quantidade de memória fisica da maquina
		for ( ; !colProps.atEnd(); colProps.moveNext()) { 
			p = colProps.item();
			//converte o valor para megabytes
			totalMemory += ( p.Capacity/1048576 );
		}
	//imprime o total de memória fisica instalada na maquina
	WScript.Echo ("memoria total: "+totalMemory+" mb");
//retorna o total de memória fisica
return totalMemory;
}

//--------------------------------------------------------------------------

//Função 5 Serviços rodando na máquina

var servicos = function (){
//obtem objeto que representa o servico WMI da máquina	
var objWMIService = GetObject( "winmgmts://./root/cimv2" );
//obtem a lista de serviços instalados no sistema
var colItems = objWMIService.ExecQuery("Select * from Win32_Service	",null,48);
var colProps = new Enumerator(colItems);
//variavel para armazenar a lista de serviços
var serviceArray = new Array ();
//le a lista de serviços e armazena na variavel
for ( ; !colProps.atEnd(); colProps.moveNext()) { 
	p = colProps.item();
	var obj = new Object ();
	serviceArray.push (p);
}

//le cada elemento da variavel array e imprime as informações do serviço
for (var i = 0; i< serviceArray.length; i+=1) {
	var service = serviceArray[i];
	//WScript.Echo ("nome: "+service.Name );
	//WScript.Echo ("nome fantasia: "+service.DisplayName );
	//WScript.Echo ("status: "+service.State  );
	//WScript.Echo  ();
}
//retorna o array de serviços
return serviceArray;
}

//--------------------------------------------------------------------------
//Função 7 serviços iniciados com o S.O

var processo = function (){
//obtem objeto que representa o servico WMI da máquina
var objWMIService = GetObject( "winmgmts://./root/cimv2" );
//obtem informações sobre os processos na lista inicialização
var colItems = objWMIService.ExecQuery("Select * from Win32_StartupCommand	",null,48);
var colProps = new Enumerator(colItems);
//variavel para armazenar a lista de processos
var processArray = new Array ();
//le a lista de processos e armazena na variavel
for ( ; !colProps.atEnd(); colProps.moveNext()) { 
	p = colProps.item();
	var obj = new Object ();
	processArray.push (p);
}

//le cada elemento da variavel array e imprime as o nome dos processos
for (var i = 0; i< processArray.length; i+=1) {
	var process = processArray[i];
	//WScript.Echo ("nome: "+process.Name );
}
//retorna o array de processos
return processArray;
}


//--------------------------------------------------------------------------\\

//Início do relatório

WScript.Echo("\n");
WScript.Echo("Relatório de Manutenção");
WScript.Echo("\n");


//Função 1
//Imprime o nome do computador
WScript.Echo("Nome do computador: " + nomepc());


//Função 2
//Recebe o resultado da função tipoeq verifica se é verdadeira (Laptop) ou Falsa (Desktop)
if (tipoeq())
	WScript.Echo ("O equipamento eh laptop");
	else	
	WScript.Echo ("O equipamento eh desktop");


//Função 3
//Recebe o resultado da função tipoeq verifica se é verdadeira (Laptop) ou Falsa (Desktop)
if(tipoeq()){
//Se leptop: verifica quantidade de memomoria (mínimo 2Gb)
	if (qtdmemoria() < 2048)
	WScript.Echo("Equipamento fora das especificações da empresa");
	else
	WScript.Echo("Equipamento de acordo com as especificações da empresa");
}
//Se desktop: verifica quantidade de memomoria (mínimo 4Gb)
else {
if (qtdmemoria() < 4096)
	WScript.Echo("Equipamento fora das especificações da empresa");
	else
	WScript.Echo("Equipamento de acordo com as especificações da empresa");
}
WScript.Echo("\n");


//Função 5 - WinDefend, sppsvc, MpsSvc
//variavel está recebendo a array da função
var varservice = servicos();

//o laço de repetição está lendo o array em busca dos serviços
for (var i = 0; i< varservice.length; i+=1) {
		var varservice2 = varservice[i];
//Caso os serviços WinDefend, sppsvc, MpsSvc não estejam ativos informa ao usuário
	if(varservice2.Name=="WinDefend"){
		if (varservice2.State=="Stopped")
		WScript.Echo("ALERTA DE SEGURANÇA! O serviço WinDefend não está rodando.");
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
WScript.Echo("\n");

//Função 7 Verificar OneDrive
//variavel está recebendo a função
var varprocess = processo();
//o laço de repetição está lendo o array em busca do processo
for (var i = 0; i< varprocess.length; i+=1) {
	var varprocess2 = varprocess[i];
	//Caso o OneDrive seja  encontrado na lista check=1, caso não seja encontrado check=0
	if(varprocess2.Name=="OneDrive"){
var check = 1;
break;
}
else
check =0;
}
//Caso a verificação seja verdadeira (check=1) o laço informa que o processo está na lista.
//Caso a verificação seja falsa (check=0) o laço informa que o processo não está na lista.
if (check==1)
WScript.Echo("O OneDrive está na lista de inicialização");
		else
		WScript.Echo("O OneDrive não está na lista de inicialização");

