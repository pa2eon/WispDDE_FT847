WiSP DDE Client Ver. 4.3.8
jan-2005 Fernando Mederos - CX6DD
October - 2020 Eric Oosterbaan - PA2EON

	** Please activate line-wrap to read this text **

	INTRODUCTION:
	-------------

	"WiSP DDE Client" is a Win95/98/2000(*) application to control radios and rotors according to calculations made by a separate satellite tracking program.

	Windows provide a handy way of passing information among pplications: Dynamic Data Exchange or DDE. Thanks to this, applications running concurrently can exchange any kind of information, for example the name of a satellite being tracked, its position, up and downlink frequencies and operating modes. WiSP DDE Client takes advantage of this and works as a sofware interface between the tracking program making orbit calculations and your rotors and radios.

	WiSP DDE Client accepts DDE information from WiSP, AmSAT-BDA's Station Program, EA4TX's ARS, WinOrbit, Nova for Windows, SatPC32, Satscape, WXtrack and Orbitron. Station sends only rotor positions data thru DDE so no rig control is possible in this case.

	Currently supported rotor interfaces are: GS-232, FODTrack, CI-V, IF-100 and TrakBox. CI-V rotor interface is one that connects to the same CT-17 box as every Icom rig, so a single serial port can be used to drive rotors and radios.

	Radios supported are: IC-821, IC-970, IC-275, IC-475, IC-746, IC-R7000, IC-R8500, PCR-1000, FT-847, FT-817, FT-100, FT-736, VR-5000, AR-5000, AR-8000, TH-D7, TM-D700, TS-790 and TS-2000.



	(*) Note: to use parallel ports under WinNT/2000 the program needs PORT95NT to be installed. PORT95NT.ZIP file in WiSPDDE's directory contains the installer for this program.



	Some credit is due: AR control routines and WinOrbit DDE addition thanks to David-EB5DGP. FT-817, FT-100 and TS-2000 control routines by Howard-G6LVB.


	Distrubution:

	WiSP DDE Client is distributed free of charge for Radio Amateur use.
	Latest version is always available at http://www.laboratoriomederos.com/CX6DD/wispdde


	Multi-rig feature:

	Starting with version 3.0 there is a multi-rig feature which enables WiSPDDE to pick two radios from a list to take care of downlink channel and uplink channel respectively. Both channels may be attended by the same radio provided it is dual-band. Information such as tracked satellite, mode and frequency will be taken into account during the automatic selection process which is performed each time the tracked satellite changes. There is great flexibility in the selection combinations: the radios need not be connected to the same port nor they have to be of the same brand etc. Again, a single dual-band radio may be selected for both uplink and downlink channels.


	Transparent-tuning feature:

	Starting with version 4.0, WiSPDDE supports what has lately been called "Transparent Tuning" in order to make it suitable for analog transponder operation.


	INSTALLATION:
	-------------

	Starting with version 4.1, WiSPDDE is distributed with a standard 'setup.exe' program which extracts and installs all needed files including shared libraries. You can select a directory path on your hard-drive for the program to be copied and an entry-name in the Start menu for a shortcut to execute the program.

	As of february-2004 SatPC32 has been modified to work as a DDE-Server instead of DDE-Client, so no special configuration is needed any more in order to let WiSPDDE get tracking info from this program.


	CONFIGURATION:
	--------------

	When first run WiSP DDE Client is setup to receive DDE from WiSP and to drive no Radio nor Rotor interfaces.
	If you are replacing an older version of WiSPDDE errors may be reported during program start-up preventing any further operation. This is due to incompatibilities in the way that different versions store configuration information. To solve this you can eliminate all previous configuration and start from scratch configuring the new version of WiSPDDE. Execute the program with an R argument in the command-line: put "wispdde R" in Execute... option in Start Menu.


	DDE Link:

	You should start by making sure that DDE selection is correct in the DDE Link settings window and the refresh rate suits your needs. Too fast a rate may prevent the control cycle from being completed, a minimum interval of 1 second is reccomended.
	The Decimal Separator must also be configured according to the numeric format used by the orbital-prediction program ('.' or ','). The Dec. Sep. character can be found in the Az. and El. values of WiSPDDE window.


	Radio:

	Then go to the Radio menu item and select the radio number you wish to configure from the Radio Index entry. You will have to check the Enable checkbox to gain access to the rest of the entry fields. There is no limit over the number of rigs you can add to the list, under some circumstances it may be useful to give more than one Index number to the same rig.
	Then specify the model of the radio, the baudrate and the port it is attached to and its CI-V address (only aplicable to Icom radios). Note that PCR-1000 baudrate is forced at 9600.
	Bidirectional Interface checkbox is only relevant for Icom rigs and it tells whether or not to expect any reply from the rig after a command is sent. This is helpful for using simple home-made CT-17-like iterfaces that do not provide a data path back from the rig to the computer.
	TNC up/dn for SSB downlink checkbox is included to let the TNC with a PSK modem make corrections over the downlink frequency after the coarse frequency is set by the computer. If this is checked the downlink frequency will only be set a few times after the pass has begun and set free after that. Appart from this, while the satellite keeps below 3deg. of elevation, downlink frequency will still be updated. This feature will only take effect when downlink mode is USB or LSB, in any other mode operation will be normal.

	FT-847 CTCSS Activation:

	In version 4.3.8 added the CTCSS code for use with the FT-847 radio.
	The original VB code is used with the 'old' VB compiler.
	vy '73 PA2EON
	

	Auto-Selection:

	The Auto-Selection Config. button gives access to the set of tracking conditions imposed for the automatic selection of the current radio.
	If you want the radio to be used as downlink channel, check Downlink; to be used as uplink channel, check Uplink. If you check none the radio will never be automatically selected.
	In the Satellites field type the names of the satellites you want to be tracked by this radio, or simply "ALL". There is no limit to the number of satellites entered. Be carefull to type the name *exactly* as WiSP calls it, otherwise it will not be recognized.
	In the Modes entry field type the modes you want this radio to use or just "ALL". Again there is no limit to the number of modes entered, but WiSP currently knows about USB LSB CW CW-N FM FM-N and FM-W.
	In the Freqs. entry field type the ranges you will cover with this radio, ranges are specified with the starting MHz, a slash '-' and the ending MHz. You can enter as many ranges as you like separated by spaces and they need not be in any particular order. You may enter "ALL" in this field also.
	Accesory port control is provided to drive devices such as relays to automatically change the station's configuration according to the tracking needs. If Accesory Port is enabled, the Output Value will be sent to the specified Port Address. Downlink and Uplink radios can share the same Accesory Port or not, there is no restriction over this. If they do share a single Acc.Port, both Output Values will be ORed and sent to the Port Address.
	There is a convenient Reset button to initialize all fields with sample values to use as reference.
	Save your settings and then Close to return to Radio Settings window, then Save again to save general radio settings. You may go on to the next radio now or click Close to return to Main window.
	Remember to Save your changes prior to switching to a different radio Index, otherwise you will loose them.
	During normal program execution an automatic-selection process begins each time the tracked satellite changes. This process consists of checking the set of conditions stored for each configured radio starting with radio number 1. If the tracked satellite name or frequency or mode is out of the scope of this radio it will be discarded and the next radio in the list will be checked until a suitable one is found. This is done twice, once for the downlink channel and then again for the uplink channel.


	Rotor:

	Then go to the Rotor menu item and the rotor interface settings window will show. Begin by choosing the interface type you will use.
	Bidirectional Interface checkbox is the same as the above explained for Radio settings. The current CI-V interface does not acknowledge the computers' commands, so this will remain unchecked in case you use this type of interface.
	Step size is used to round the actual rotor values received from WiSP to the nearest integer multiple of the number entered.
	An Auto Flip detection is provided because old versions of WiSP did not send flipped data over DDE even if they were setup to do so. Newer releases (starting with GSC 2.03 I believe) does send flipped rotor positions. If this option is checked, a simple algorithm is used to detect the posibility of a satellite getting into the stop position of the azimuth rotor. This is acomplished by examining where did the satellite come from and where is it going to. Assume North rotor-stop position; if the sat appeared from the West and Azimuth is increasing then it flips, if the sat appeared from the East and Azimuth is decreasing it flips, in any other case it does not flip.
	South rotor-stop position selection is available in case it is needed.
	Save your settings and then return to main window with Close button.

	NOTE: Complete information regarding the CX6DD rotor control interface in its CI-V and GS-232 flavours is available from http://www.laboratoriomederos.com/cx6dd


	USAGE of WiSPDDE with DDE Servers:
	----------------------------------

	In order for this program to be able to interface a DDE Server (tracking program) with rotors and radios it must be running concurrently with it (minimized is OK). It is not relevant wether you start it before or after the server, but sats will not be tracked until the server is running too. DDE interface type must be selected in the configuration section of the tracking program to make it start sending info thru DDE. 


	MISCELANEOUS:
	-------------

	Automatic-Update checkboxes are provided in the Main Window to let manual changes be made over downlink and uplink radios and rotor interfaces without interference.
	The Radio and Rotor buttons are provided to instruct the program to immediately send the Radios or Rotor data respectively to the proper hardware interface.

	The program can be run in a Single-Pass mode by specifying an "s" argument in the command-line (wispdde.exe s). In this mode, WiSPDDE will wait for a satellite pass and as soon as it is over will close itself. This is very handy to run WiSPDDE from a WiSP schedule event asociated to all satellites.


	TO-DO:
	------

	Comprehensive documentation!
	There are features that need documentation in this text, some are evident, some are hidden. In general I provided a short tool-tip explanation on every important control of the program. Please read them. You just need to place the mouse pointer some seconds over the control to see the tool-tip text appear.


	I hope you find this program useful. 
	Comments and bugs: CX6DD@amsat.org.

	73 de Fernando, CX6DD.


Fernando Mederos, jan-2005.
