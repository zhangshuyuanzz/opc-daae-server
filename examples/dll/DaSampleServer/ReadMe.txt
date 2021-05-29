-------------------------------------------------------------------------
OPC DA/AE Server SDK DLL                Base DA/AE Sample Server Adaption
-------------------------------------------------------------------------
Copyright (c) 2011-2021 Technosoftware GmbH. All rights reserved

This C++ based plugin for the OPC DA/AE Server SDK DLL shows a base 
OPC DA 2.05a/3.00 and AE 1.00/1.10 server implementation.
  
At startup items with several data types and access rights are statically 
defined. 

The RefreshThread simulates signal changes for the items
	SimulatedData.Ramp
	SimulatedData.Random
	SimulatedData.Sine
and writes the changed values into the internal cache and the generic 
server cache.

Item values written by a client are written into the local buffer only.

Files in this sample:
- ClassicNodeManager.h / ClassicNodeManager.cpp
    The node manager base on the interface IClassicBaseNodeManager and 
    contains the methods that are called by the generic server and need 
    to be implemented application specific. 
- IClassicBaseNodeManager.h / IClassicBaseNodeManager.cpp
    Defines the generic server interface. DON'T CHANGE THIS FILE.
    It contains definitions, callback methods and default implementations 
    of the methods call by the generic server.

- OpcDllDaServer.exe
    This is the generic OPC DA 2.05a/3.00 server

Post Build Steps
After a successful compilation the following steps are executed in the post 
build event:
1. copy the files OpcDllDaServer.exe from the project directory 
   into the bin directory
2. Register the server by executing OpcDllDaServer.exe -regserver
				
The server is now registered and can be accessed by OPC DA 1.0a/2.05a/3.00 
clients.

Debugging
To debug the plug-in assembly you need to:
1.  Open project properties and select  
		Configuration Properties  -  Debugging
		Select 'Start external program' and browse to OpcDllDaServer.exe 
		in bin
2.  Set Breakpoints
3.  Start the program execution

The generic server is initialized and calls the plug-in methods 

	OnGetDaServerRegistryDefinition()
	OnGetDaServerParameters() 
	OnCreateServerItems() 
	
Then the server is idle until a client connects.

The further activity depends on the plug-in and the client access.


