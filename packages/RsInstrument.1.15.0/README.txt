RsInstrument is a .NET component that provides convenient way of communicating with R&S instruments.

Check out the full documentation here: https://rsinstrumentcsharp.readthedocs.io/

Examples: https://github.com/Rohde-Schwarz/Examples/tree/main/GeneralExamples/Csharp/RsInstrument

Version history

Version 1.15.0.67 - 21.10.2021
	- Added .NET Standard 2.0 allowing targeting .NET Core and .NET Framework with one assembly.
	- Added RohdeSchwarz.RsInstrument.Conversions namespace with double,integer,boolean conversion extention methods.

Version 1.14.0.65 - 15.10.2021
	- Fixed CheckStatus() which was skipped if the QueryInstrumentStatus was false. Now the error checking is performed regardless that settings
	- Added correct conversion of strings with SI suffixes (e.g.: MHz, KHz, THz, GHz, ms) to double, int32, int64
	- Fixed VISA read buffer in case of multi-threading access

Version 1.13.0.64 - 28.09.2021
	- Fixed bug where the NuGet packages contained debug versions of the assemblies with file version 1.0.0.0
	- Additional changes only relevant to auto-generated drivers

Version 1.11.0.61 - 19.05.2021

	- Added constructor RsInstrument(string resourceName, string optionString)
	- improved options string help
	- added checking for empty or null resourceName in the constructor

Version 1.10.1.60 - 18.04.2021
	
	- Added alias methods:
	
		- Query() = QueryString()
		- Write() = WriteString()
		- QueryWithOpc() = QueryStringWithOpc()
		- WriteWithOpc() = WriteStringWithOpc()

Version 1.10.0.57 - 19.01.2021

	- Added documentation on https://rsinstrumentcsharp.readthedocs.io/
	- Changes relevant to auto-generated drivers only

Version 1.9.0.56 - 14.01.2021

	- Added QueryOpc(int visaTimeout)
	- Fixed error where the System.TimeoutException was thrown instead of the RsInstrument.VisaTimeoutException
	- Cosmetic changes

Version 1.8.0.55 - 14.12.2020

	- Fixed setting of VISA Timeout by init to 10000ms
	- Added "DTX", "Dtx", "dtx" to a list of values that are represented as NaN

Version 1.7.3.53 - 25.11.2020

	- NuGet package signed with Rohde Schwarz certificate
	- Core change: Only relevant for auto-generated instrument drivers

Version 1.7.2.51 - 16.11.2020

	- Changed NuGet icon
	- Adjusted Company name and copyright
	- Core change: Only relevant for auto-generated instrument drivers

Version 1.7.0.50 - 11.11.2020

	- Changed authors and copyright information
	- Core change: Conversion of the empty returned string to array returns empty array. Before, the empty string was converted to an array of one empty element.
	- Added QueryStringList(), QueryStringListWithOpc()
	- Added QueryBooleanList(), QueryBooleanListWithOpc()

Version 1.6.4.48 - 09.11.2020

   - Fixed parsing of the instrument errors when an error message contains two double quotes

Version 1.6.3.47 - 22.10.2020

   - Changes only relevant for auto-generated instrument drivers
   - Added 'UND' to the list of numbers that are represented as NaN

Version 1.6.0.43 - 05.10.2020

    - New Core with added OptionsString token 'TermChar' for setting a custom termination character
    - Added 'Hameg' to the list of supported instruments
    - Added static method AssertMinVersion() for checking the RsInstrument minimum version

Version 1.5.2.42 - 17.09.2020

    - Changes only relevant for auto-generated instrument drivers

Version 1.5.1.41 - 04.09.2020

	- New Core 1.8.2.41 with the fix for instrument that do not support OPT? query

Version 1.5.1.40 - 24.08.2020

	- New Core 1.8.1.40 with the fixed simulation mode issues

Version 1.5.0.39 - 11.08.2020

	- Multi-target frameworks .NET Standard 2.1, .NET Core 3.1, .NET Framework 4.5 and 4.8
	- New Core 1.8.0.38 with these features:
	- Implemented SocketIO Visa Plugin that does not need VISA
	- New Options token: 'SelectVisa' with parameters: NativeVisa | RsVisa | RsVisaPrio | Socket
	- Options token 'PreferRsVisa' is now obsolete (but still supported)
	- Added new static function RsInstrument.FindResources()
	
Version 1.4.2.38 - 04.08.2020

	- Fixed buffer size for Nrp-Z sessions
	- Added and corrected examples

Version 1.4.0.36
	- Distributed as NuGet package
	
	- Changed Core to allow for AnyCPU build
	- Added Session Settings bool AssureResponseEndWithLF

Version 1.3.0.34 - 19.06.2020

	- Added invoking read_segmented event for the first chunk of the ReadUnknownLength()
	- New Core with RepeatedCapabilities for command groups
	

Version 1.2.0.32 - 10.01.2020

	- New Core with reworked session settings
	- Support for NRP-Zxx instruments

Version 1.1.0.30 - 29.11.2019

	- Reorganized Utilities interface to sub-groups
	- Added Write/Query With Opc Event
	- Added locking for multithreading safety
	- Added segmented read / write events

Version 1.0.0.20

	- First released version