/*
5 Dez. 2015

Ricardo Akira Paiva Ichikawa
2o Semestre de Redes de Computadores
ADO Final de Algoritmos de Programação
Professor Fábio de Toledo Pereira

----------------------------------------------------------------------------------------------------------------------------

Função abaixo não está relacionada ao trabalho.
Repete a impressão de texto e quebra de linha baseada em valores recebidos
Padrão: 1x Echo (string + \n) */

var printit = function (a, b, c){
	do {
		WScript.Stdout.Write (a);
		c--;	

		var d = b;
		
		if (d == 0)
			null;
		else {
			do {			
				WScript.Stdout.Write ("\n");
				d--;
			} while (d > 0);
		}
	} while (c > 0);
}

//O conteudo da ADO inicia aqui
//A linguagem está em ingles para facilitar a escrita


//BEGINNING OF FUNCTIONS

/*-------------------------------------------------------------------------------------------------------------------------------------
1. Função para obter o nome do computador. Imprimir no relatório.*/

var getpcname = function (){

	//getss the object that represents the WMI service of the machine
	var objWMIService = GetObject ("winmgmts://./root/cimv2");

	//getss the list of windows systems in the machine
	var colItems = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem",null ,48);

	//getss the colection iterator
	var colProps = new Enumerator (colItems);

	//creates a variable that will receive the computer name
	var pcName;

		//iterates the colection of data from WMI service
		for ( ; !colProps.atEnd(); colProps.moveNext()) { 
			//getss a item from the colection
			p = colProps.item();
			//extract the computer name from the item
			pcName = p.name;
		}
	//return the computer name to the function
	return pcName;
}

/*-------------------------------------------------------------------------------------------------------------------------------------
2. Função para saber se o computador é do tipo laptop ou desktop. Imprimir no relatório.*/

var getpctype = function (){

	var myComputer = ".";
	//gets the object that represents the WMI service of the machine
	var objWMIService = GetObject ( "winmgmts:\\\\" + myComputer + "\\root\\cimv2" );

	//getss battery status
	var colItems = objWMIService.ExecQuery ("Select * from Win32_Battery");

	//sets the machine as desktop. If there is no alteration after the iteraction, the answer will be "Desktop"
	var IsLaptop = false;

	//gets the item iterator
	var objItem = new Enumerator (colItems);

		//iterates the item. If found, sets the variable as "True", saying that it is a Laptop
		for (; !objItem.atEnd(); objItem.moveNext()) {
			IsLaptop = true;
		}
/*	
		//function set string (Laptop or Desktop)
		if (IsLaptop)
			IsLaptop = "Laptop";
		else
			IsLaptop = "Desktop";
*/

	return IsLaptop;
}

/*
        .---.
       /o   o\
    __(=  "  =)__
     //\'-=-'/\\
        )   (_
       /      `"=-._
      /       \     ``"=.
     /  /   \  \         `=..--.
 ___/  /     \  \___      _,  , `\
`-----' `""""`'-----``"""`  \  \_/
                             `-`                            
FOCA NO CÓDIGO
*/

/*-------------------------------------------------------------------------------------------------------------------------------------
4. Função que retorna a % de uso do HD. Imprimir % de espaço livre. Se espaço livre > 10%, alertar para fazer limpeza de disco.*/

var gethdinfo = function (){

	//gets the object that represents the WMI service of the machine
	var objWMIService = GetObject ("winmgmts://./root/cimv2");

	//getss status of non-removable storage units
	var colItems = objWMIService.ExecQuery ("Select * from Win32_LogicalDisk where DriveType=3", null, 48);

	//gets the item iterator
	var colProps = new Enumerator (colItems);

		WScript.Echo ("\n\nStorage:");

		//iterates the data of each storage unit, returning the usage status
		var diskunit = []; var arrayguide = 0; //this variable stores each unity free space %
		for (var count = 1; !colProps.atEnd(); colProps.moveNext()) {

			p = colProps.item();

			//prints unity list number
			WScript.Echo("   Unit " + count + ":");

			//prints unity letter
			WScript.Echo("      Name: " + p.name);

			//prints unity total size
			WScript.Echo("      Size: " + (p.Size/(1024*1024*1024)).toFixed(2));

			//prints unity free space
			//WScript.Echo("      Free space: " + (p.FreeSpace/(1024*1024*1024)).toFixed(2));

			//prints disk usage %
			WScript.Echo("      Usage: " + ((1 - p.FreeSpace/p.Size)*100).toFixed(2)+"%\n");

			//storing the current unity free space %
			diskunit[arrayguide] = (100 - ((1 - p.FreeSpace/p.Size)*100)).toFixed(2);

			count++; //just a display count to list the storage units (display list starts from 1)
			arrayguide++; //counter to 'diskunit' variable (starts from 0)
		}

	//returns array with free space of each unity
	return diskunit;
}

/*-------------------------------------------------------------------------------------------------------------------------------------
5. Função que retorna serviços instalados no computador (nome, nome fantasia e status). 
Imprimir se os serviços a seguir estão rodando (status: Running): WinDefend, sppsvc, MpsSvc. Se não, reportar uma brecha de segurança.*/

var getpcservices = function (){

	//gets the object that represents the WMI service of the machine
	var objWMIService = GetObject ("winmgmts://./root/cimv2");

	//gets the list of windows services
	var colItems = objWMIService.ExecQuery ("Select * from Win32_Service", null, 48);

	//gets the item iterator
	var colProps = new Enumerator(colItems);

	//creating the array to store the service list
	var serviceArray = [];

		//storing the service list in the array
		for (; !colProps.atEnd(); colProps.moveNext()) { 

			p = colProps.item();
			var obj = new Object ();
			serviceArray.push (p);
		}

/*
		WScript.Echo ("\n\nServices:");

		//lists the services in the array
		for (var i = 0; i < serviceArray.length; i+=1) {

			var service = serviceArray[i];

				WScript.Echo("   Name: " + service.Name);
				WScript.Echo("      Display name: " + service.DisplayName);
				WScript.Echo("      Status: " + service.State + "\n");
*/

	return serviceArray;
}

/*-------------------------------------------------------------------------------------------------------------------------------------
7. Função que retorna processos programados para iniciar com o windows. Verificar se OneDrive esta na lista e se esta ativo.*/

var startuptasks = function (){

	//gets the object that represents the WMI service of the machine
	var objWMIService = GetObject( "winmgmts://./root/cimv2" )

	//gets the list of windows startup tasks
	var colItems = objWMIService.ExecQuery("Select * from Win32_StartupCommand	",null,48)

	//gets the item iterator and creating the array to store the task list
	var colProps = new Enumerator(colItems);
	var processArray = [];

		//storing the task list in the array
		for ( ; !colProps.atEnd(); colProps.moveNext()) { 
			p = colProps.item();
			var obj = new Object ()
			processArray.push (p);
		}
/*
		WScript.Echo ("\n\nStartup Tasks:");

		//lists the services in the array
		for (var i = 0; i< processArray.length; i+=1) {
			var process = processArray[i];
			WScript.Echo ("   Name: "+process.Name );
		}
*/

	return processArray;
}

//END OF FUNCTIONS

//-------------------------------------------------------------------------------------------------------------------------------------

//BEGINNING OF THE REPORT

printit("-",0, 50);printit("\n");
printit("Computer maintenance report.",2);
printit("-",0, 50);printit("\n");

//1. prints the computer name
printit("Name: " + getpcname());

//2. prints the computer type (desktop or laptop)
if(getpctype())
	printit("Type: Laptop")
else
	printit("Type: Desktop")

//4. prints warning if disk space is below 10%
var checkspace = gethdinfo();
	var arrayguide = 0; //temporary counter
	var count; //temporary counter

	for (count = 1; count <= checkspace.length; count++){

		//prints the free space of each unity
		printit("   Unit " + count + " free space: " + checkspace [arrayguide] + "%");

		//check free space on all disks
		if (checkspace [arrayguide] < 10)
			printit("      WARNING: Unit " + count + " free disk space is below 10%. Disk cleanup is recommended.",2);

		arrayguide++;
	}

//5. prints if services are running (WinDefend, sppsvc, MpsSvc)

printit("\nService check:");
var getpcservices2 = getpcservices();

	//loop for verify the services
	for(var count = 0;count < getpcservices2.length; count++){

		var checkservices = getpcservices2[count];
		//verify if MpsSvc is running
		if (checkservices.Name == "MpsSvc"){
			if (checkservices.State == "Stopped")
				printit("   !!!!!!!SECURITY WARNING: The service MpsSvc isn't running!!!!!!!")
		}

		//verify if sppsvc is running
		if (checkservices.Name == "sppsvc"){
			if (checkservices.State == "Stopped")
				printit("   !!!!!!!SECURITY WARNING: The service sppsvc isn't running!!!!!!!")
		}

		//verify if WinDefend is running
		if (checkservices.Name == "WinDefend"){
			if (checkservices.State == "Stopped")
				printit("   !!!!!!!SECURITY WARNING: The service WinDefend isn't running!!!!!!!")
		}
	}

//7. prints startup task schedule
printit("\nStartup process check:");
var startuptasks2 = startuptasks();
	for(var count = 0;count < startuptasks2.length; count++){

		var onedrivetrue; 
		var checkprocess = startuptasks2[count];

		//verify if OneDrive is in the startup list (set onedrivetrue to 1) or not on the list (set to 0)
		if (checkprocess.Name == "OneDrive"){
			onedrivetrue = 1;
			break;
		}
		else
			onedrivetrue = 0;
	}

	if (onedrivetrue == 1)
		printit("   OneDrive is in the startup list.",2);
	else
		printit("   OneDrive isn't in the startup list.",2);

printit("-",0, 50);printit("\n");
printit("End of maintenance report.",2);
printit("-",0, 50);printit("\n");


//END OF THE REPORT

printit("Ricardo Akira Paiva Ichikawa\n2o Semestre de Redes de Computadores\nADO Final de Algoritmos de Programa\u00e7\u00e3o\nProfessor F\u00e1bio de Toledo Pereira",3);
