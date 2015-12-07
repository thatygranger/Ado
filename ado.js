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

		d = b;
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

	//gets the object that represents the WMI service of the machine
	var objWMIService = GetObject ("winmgmts://./root/cimv2");

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
			
		//function return string (Laptop or Desktop)
		if (IsLaptop)
			IsLaptop = "Laptop";
		else
			IsLaptop = "Desktop";

	return IsLaptop;
}

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
		var diskunit = new Array (); var arrayguide = 0; //disk cleanup warning variables
		for (var count = 1; !colProps.atEnd(); colProps.moveNext()) {

			p = colProps.item();

			WScript.Echo("   Unit " + count + ":");
			WScript.Echo("      Name: " + p.name);
			WScript.Echo("      Size: " + (p.Size/(1024*1024*1024)));
			//WScript.Echo("      Free space: " + (p.FreeSpace/(1024*1024*1024)).toFixed(2));
			WScript.Echo("      Usage: " + ((1 - p.FreeSpace/p.Size)*100)+"%\n");

			diskunit[arrayguide] = (100 - ((1 - p.FreeSpace/p.Size)*100));

			count++; 
			arrayguide++;
		}	

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
	var serviceArray = new Array ();

		//storing the service list in the array
		for (; !colProps.atEnd(); colProps.moveNext()) { 

			p = colProps.item();
			var obj = new Object ();
			serviceArray.push (p);
		}

		WScript.Echo ("\n\nServices:");

		//lists the services in the array
		var checkservice = new Array (); var count = 0; //variables to verify the specified services
		for (var i = 0; i < serviceArray.length; i+=1) {

			var service = serviceArray[i];

				WScript.Echo("   Name: " + service.Name);
				WScript.Echo("      Display name: " + service.DisplayName);
				WScript.Echo("      Status: " + service.State + "\n");

					//verify if 'WinDefend' is running
					if (service.Name == "MpsSvc"){

						//printit("^^^^^^^^^^^^^^^^^",10,2); //mark the service on the list

						if (service.State == "Running")
							checkservice[count] = 1; //running == 1, array [0]
						else 
							checkservice[count] = 0; //stopped == 0, array [0]

						count++;
					}

					//verify if 'sppsvc' is running
					if (service.Name == "sppsvc"){

						//printit("^^^^^^^^^^^^^^^^^",10,3); //mark the service on the list

						if (service.State == "Running")
							checkservice[count] = 1; //running == 1, array [1]
						else 
							checkservice[count] = 0; //stopped == 0, array [1]
						
						count++;
					}
					//verify if 'MpsSvc' is running		
					if (service.Name == "WinDefend"){

						//printit("^^^^^^^^^^^^^^^^^",10,2); //mark the service on the list

						if (service.State == "Running")
							checkservice[count] = 1; //running == 1, array [1]
						else 
							checkservice[count] = 0; //stopped == 0, array [1]
						
						count++;
					}
		}

	return checkservice;
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
	var processArray = new Array ();

		//storing the task list in the array
		for ( ; !colProps.atEnd(); colProps.moveNext()) { 
			p = colProps.item();
			var obj = new Object ()
			processArray.push (p);
		}

		WScript.Echo ("\n\nStartup Tasks:");
		//lists the services in the array
		var checkonedrive = 0;
		for (var i = 0; i< processArray.length; i+=1) {
			var process = processArray[i];
			WScript.Echo ("   Name: "+process.Name );

			//verify if 'OneDrive' task is in the startup list 		
			if (process.Name == "WinDefend"){
				checkonedrive = 1;
			}
		}
	return checkonedrive;
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
printit("Type: " + getpctype());

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

//5. prints if services are running

var checkservices = getpcservices();
var thoseservices = new Array("MpsSvc", "sppsvc", "WinDefend");

	printit(); printit("!",0,62); printit("\n");

	//verify if MpsSvc, sppsvc and WinDefend are up or down
	for (count = 0; count <= checkservices.length; count ++){
		if (checkservices[count] == 0)
			printit("      SECURITY WARNING: Service '" + thoseservices[count] + "' is disabled."); //this warning will be printed at the end of the service list
	}

	printit(); printit("!",0,62); printit("\n");

//7. prints startup task schedule
var checkonedrivestartup = startuptasks();
	//check if OneDrive is on the list
	if (checkonedrivestartup == 1)
		printit("\nOneDrive is in the startup list and will start automatically after logon.",3);
	else
		printit("\nOne drive isn't in the startup list.",3);


printit("-",0, 50);printit("\n");
printit("End of maintenance report.",2);
printit("-",0, 50);printit("\n");

printit("Ricardo Akira Paiva Ichikawa\n2o Semestre de Redes de Computadores\nADO Final de Algoritmos de Programa\u00e7\u00e3o\nProfessor F\u00e1bio de Toledo Pereira",3);
