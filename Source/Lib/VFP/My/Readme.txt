Visual FoxPro "Sedna"
=====================



Sedna is a collection of libraries, samples and add-ons to Visual FoxPro 9.0 SP2.

Please uninstall previous CTPs or Beta from your machine before installing Sedna. By default Sedna Setup will install the files onto your disk under the "Microsoft Visual FoxPro 9\Sedna" folder.

This installation contains six components: VistaDialogs4COM, Upsizing Wizard, Data Explorer, NET4COM and MY for VFP and DDEX for VFP. The following sections describe these components and provide guidance on how to get started.




VistaDialogs4COM
================

The VistaDialogs4COM is a collection of COM-visible classes that wrap the functionality provided by the Microsoft VistaBridgeLibrary. VistaDialogs4COM provides Visual FoxPro developers access to the Windows Vista TaskDialog and Common Dialogs.

Setup will install the VistaDialogs4COM assembly, the VistaBridgeLibrary and other required DLLs, the VistaDialogs4COM source. A sample VFP project demonstrating use of different features of VistaDialogs4COM is also included.

The VistaDialogs4COM folder contains the following:

VistaDialogs4COM.dll	-- DLL for the COM components that wrap VistaBridgeLibrary.
VistaDialogs4COM	-- Folder containing the VB.NET source code for VistaDialogs4COM
VFP Sample		-- Folder containing a VFP sample project that illustrates the use VistaDialogs4COM

A few additional notes on VistaDialogs4COM:

VistaLibrary4COM requires Windows Vista. The API used are not available on earlier versions of Windows. The VFP Sample folder contains images that show the different dialogs and how they differ from the corresponding dialogs in earlier versions.



SQL Server Upsizing Wizard
==========================

This is an update to the Visual FoxPro 9.0 SP2 Upsizing Wizard. The Sedna Upsizing Wizard will install to the "Microsoft Visual FoxPro 9\Sedna\UpsizingWizard" folder under Program Files. To launch the new wizard run the 'UpsizingWizard.app' from this location.

This update includes:

*  Updated, cleaner look and feel
*  Streamlined, simpler steps
*  Support for bulk insert to improve performance.
*  Allows you to specify the connection as a DBC, a DSN, one of the existing connections or a new connection string.
*  Fields using names with reserved SQL keywords are now delimited.
*  If lQuiet is set to true when calling the wizard, no UI is displayed. It uses RAISEEVENT() during the progress of the upsizing so the caller can show progress.
*  Performance improvement when upsizing to Microsoft SQL Server 2005.
*  Trims all Character fields being upsized to Varchar. 
*  BlankDateValue property available. It specifies that blank dates should be upsized as nulls. (Old behavior was to set them to 01/01/1900).
*  Support for an extension object. This allows developers to hook into every step of the upsizing process and change the behavior. Another way is to subclass the engine.
*  Support for table names with spaces.
*  UpsizingWizard.APP can be started with default settings (via params) for source name and path, target db, and a Boolean indicating if the target database is to be created.



Database Explorer
=================
 
This is an update to the Visual FoxPro 9.0 SP2 Data Explorer. Setup will install the Database Explorer to the "Microsoft Visual FoxPro 9\Sedna\DataExplorer" folder under Program Files. To launch the new wizard run the 'DataExplorer.app' from this location.

This update includes:

*  Fixed drag and drop of VFP table from Data Explorer to a form.
*  Fixed issue: "Drag/drop of VFP table/view from the Data Explorer to a form does not set the RecordSource."
*  Fixed issue with free tables not showing their columns when expanding the node.
*  Drag and drop operations now respect Field Mapping settings for SQL Server data
*  Allows sorting to apply to specific objects, not all.
*  Display SQL ShowPlan for local views 
*  Display ShowPlan information for VFP table/view queries in the query results 
*  Add Showplan parameter setting to the Options dialog, use in showplan features
*  Context menu item to launch the new UpsizingWizard.



NET4COM
=======

The NET4COM library is a collection of COM classes that wrap a subset of the .NET Framework 2.0. The .NET Framework is a rich collection of namespaces and API that provides a comprehensive set of functionality that developers can use to build applications that run on the .NET platform. While VFP does have a rich library of API, there are some features that either do not exist in the VFP libraries or are harder to use than in the Framework. NET4COM brings together a small subset of the .NET Framework — a collection of commonly used API that brings to VFP functionality that does not exist.

Setup will install the NET4COM assembly with sample files to disk. Setup will also register the COM components in NET4COM.dll. The NET4COM folder contains the following:

NET4COM.dll -- The DLL for the COM components that wrap a subset of the .NET Framework 2.0.
Source      -- Folder containing project file and source code for NET4COM
VFPSamples  -- Folder containing VFP sample code that uses NET4COM
VB6Samples  -- Folder containing VB6 sample code that uses NET4COM
FFC         -- FFC wrappers for NET4COM

The samples demonstrate use of NET4COM from Visual FoxPro or Visual Basic 6.0.



MY for VFP
==========

The MY library for Visual FoxPro is similar to NET4COM. MY is implemented natively for Visual FoxPro and like NET4COM exposes commonly used functionality in a hierarchy that is easy to discover and navigate.

Setup will install the MY library and related files in the "MY" folder under the selected install destination. To install MY into Visual FoxPro:

   *  Browse to the MY folder
   *  Run MY.APP

The MY folder contains a help file: MY.CHM that describes, setup and usage of MY and includes the API reference for the different APIs.



DDEX for VFP
============

This is the Visual Studio 2005 Extension (Data Designer EXtension) for Visual FoxPro data files. Installing the VFP DDEX Provider also installs the ADO .NET data provider wrapper over VFPOLEDB.

DDEX allows Visual Studio to work better with Visual FoxPro data sources. It exposes VFP metadata to the data designers within Visual Studio, thus enabling common design tasks like dragging and dropping tables from Server Explorer to Visual Studio designers.

Setup will copy the DDEX Provider, a registration utility and source code to "Microsoft Visual FoxPro 9.0\DDEXProvider" folder.

Before using, the DDEXProvider must be installed. NOTE that DDEX requires Visual Studio 2005 to be installed.

NOTE: RegDDEX must be run as an Adminstrator. Start the Windows Command Prompt as Administrator.

To register the DDEXProvider...

   *  Change the folder to 'DDEXProvider' under the Sedna folder
   *  Run RegDDEX

To unregister the DDEXProvider...

   *  Change the folder to 'DDEXProvider' under the Sedna folder
   *  Run RegDDEX /u

Once registered, Visual FoxPro will appear as one of the data source types that can be selected in the Connection Dialog of Visual Studio.

