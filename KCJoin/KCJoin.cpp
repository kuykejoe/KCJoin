////////// KCJoin Version 1.0.0
////////// Command line utility for joining and leaving local domains
////////// Joey@Kuykendall.Consulting
#include <Windows.h>
#include <LM.h>

#include <atlbase.h>
#include <atlstr.h>
#include <atlapp.h>

#include "SimpleOpt.h"

//WTL AppModule
CAppModule _Module;

//Command line option flags 
enum { OPT_HOST, OPT_USER, OPT_PASS, OPT_OU, OPT_DOMAIN, OPT_UNJOIN };

//Command line option flags parameters
CSimpleOpt::SOption g_Options[] = {
	{OPT_HOST, L"-host", SO_REQ_CMB},     // -host=HOSTNAME 
	{OPT_USER, L"-user", SO_REQ_CMB},     // -user=USERNAME
	{OPT_PASS, L"-pass", SO_REQ_CMB},     // -pass=PASSWORD
	{OPT_OU, L"-ou", SO_REQ_CMB},         // -ou=OU
	{OPT_DOMAIN, L"-domain", SO_REQ_CMB}, // -domain=DOMAIN
	{OPT_UNJOIN, L"-unjoin", SO_NONE},    // -unjoin
	SO_END_OF_OPTIONS
};

/*<
	FUNCTION: PrintError
	.Description
	Print the system error message for the last-error code.
	See https://docs.microsoft.com/en-us/windows/desktop/debug/retrieving-the-last-error-code

	PARAMETERS: None
	
	RETURN: None
>*/
void PrintError()
{
	//Create a void pointer to contain the error message
	LPVOID lpMessage = NULL;
	//Capture the last error
	DWORD dwError = GetLastError();

	//Use FormatMessage to locally allocate a string with the error message in the default language
	FormatMessage( FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
		NULL, dwError, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)& lpMessage, 0, NULL);

	//Print the error for the caller
	wprintf(L"Error %ld: %s", dwError, (LPWSTR)lpMessage);
		
	//Free the message 
	LocalFree(lpMessage); lpMessage = NULL;
}

/*<
	FUNCTION: Run
	.Description
	Processes the command line and handles the 'business logic' of the utility. 

	PARAMETERS:
	.Parameter: Refrence to CSimpleOpt csoArgs
	CSimpleOp object initialized with command line arguments.

	RETURN: bool
	True upon success, false if there was an error.
>*/
bool Run(CSimpleOpt& csoArgs)
{
	//Create bool that will be true (default) if joining, false if unjoining
	bool fJoin = true;
	//Create CString buffers for the input
	CString cszHost, cszUser, cszPass, cszOu, cszDomain;
	
	//See CSimpleOpt.h for useage information
	//Retrieve the next argument until there are none left
	while (csoArgs.Next())
	{
		//If the argument is valid
		if (csoArgs.LastError() == SO_SUCCESS)
		{
			//Extract argument information to respective buffers
			switch (csoArgs.OptionId())
			{
			case OPT_HOST: cszHost.Format(L"%s", csoArgs.OptionArg()); break;
			case OPT_USER: cszUser.Format(L"%s", csoArgs.OptionArg()); break;
			case OPT_PASS: cszPass.Format(L"%s", csoArgs.OptionArg()); break;
			case OPT_OU: cszOu.Format(L"%s", csoArgs.OptionArg()); break;
			case OPT_DOMAIN: cszDomain.Format(L"%s", csoArgs.OptionArg()); break;
			case OPT_UNJOIN: fJoin = false; break;
			}
		}
	}

	//If we are joining or unjoining
	if (fJoin)
	{
		//Check that the required arguments were specified
		if (cszUser.IsEmpty() || cszPass.IsEmpty() || cszOu.IsEmpty() || cszDomain.IsEmpty()) return false;

		//Perform join steps
		//Immediately exit upon error
		
		//If an alternate hostname was specified
		if (!cszHost.IsEmpty())
		{
			//Change the Physical DNS Hostname (first part of FQDN) and notify the caller of the results
			if (SetComputerNameEx(ComputerNamePhysicalDnsHostname, cszHost)) { wprintf(L"Successfully changed computer name: %s", cszHost.GetString()); }
			{
				wprintf(L"ERROR: Failed to SetComputerNameEx! Host: %s GLE: %ld", cszHost.GetString(), GetLastError());
				return false;
			}
		}
		
		//Create join option flag specifing to join the domain with a new computer account
		DWORD dwJoinOpts = NETSETUP_JOIN_DOMAIN | NETSETUP_ACCT_CREATE;
		//Specify to join with the new name, if a new one was applied
		if (!cszHost.IsEmpty()) { dwJoinOpts |= NETSETUP_JOIN_WITH_NEW_NAME; }
		//Join the domain
		NET_API_STATUS statusJoinResult = NetJoinDomain(NULL, cszDomain, cszOu, cszUser, cszPass, dwJoinOpts);
		//Notify the caller of the result
		if (statusJoinResult == NERR_Success) { wprintf(L"Joined %s domain using %s in container %s ", cszDomain.GetString(), cszUser.GetString(), cszOu.GetString()); return true; }
		else { PrintError(); }
	}
	else
	{
		//Check that the required arguments were specified
		if (cszUser.IsEmpty() || cszPass.IsEmpty()) return false;
		
		//Perform unjoin steps
		//Immediately exit upon error

		//If an alternate hostname was specified
		if (!cszHost.IsEmpty())
		{
			//Change the Physical DNS Hostname (first part of FQDN) and notify the caller of the results
			if (SetComputerNameEx(ComputerNamePhysicalDnsHostname, cszHost)) { wprintf(L"Successfully changed computer name: %s", cszHost.GetString()); }
			{
				wprintf(L"ERROR: Failed to SetComputerNameEx! Host: %s GLE: %ld", cszHost.GetString(), GetLastError());
				return false;
			}
		}

		//Unjoin the domain, deleting the computer account
		NET_API_STATUS result = NetUnjoinDomain(NULL, cszUser, cszPass, NETSETUP_ACCT_DELETE);
		//Notify the caller of the result
		if (result == NERR_Success) { wprintf(L"Unjoined domain using %s", cszUser.GetString()); return true; }
		else { PrintError(); }
	}
	
	return true;
}

/*<
	FUNCTION: main
	.Description
	Module entry point. Initializes, retrieves the program arguments, and calls the Run function.

	PARAMETERS: None

	RETURN: int
	0 upon success, 1 if there was an error.
>*/
int main()
{
	//Create integers to store return value and argument count
	int iReturn = 0, iArgC = 0;
	//Create a pointer to an array of strings that will contain the argument variables
	LPWSTR* pwszArgV = NULL;
	//Create an HRESULT to store return values
	HRESULT  hrLast;
	
	//Initialize COM
	hrLast = CoInitializeEx(0, COINIT_MULTITHREADED);
	//If it didn't succeed, inform the user and exit
	if (FAILED(hrLast))
	{
		wprintf(L"Failed to initialize COM subsystem! HRESULT: %x", hrLast);
		return 1;
	}

	//Initialize WTL program module
	hrLast = _Module.Init(NULL, GetModuleHandle(NULL), &LIBID_ATLLib);
	//If it didn't succeed, inform the suer and exit
	if (FAILED(hrLast))
	{
		wprintf(L"Failed to initialize WTL! HRESULT: %x", hrLast);
		//Terminate WTL module
		_Module.Term();
		//Unload COM
		CoUninitialize();
		return 1;
	}

	//Uses the shell to locally allocate an array with the command line arguments, and store count in iArgC
	pwszArgV = CommandLineToArgvW(GetCommandLine(), &iArgC);
	//If it didn't get allocated, inform the user and go to exit
	if (pwszArgV == NULL)
	{
		wprintf(L"Failed to process the command line! GLE: %ld", GetLastError());
		//Terminate WTL module
		_Module.Term();
		//Unload COM
		CoUninitialize();
		return 1;
	}

	//Create and initialize a CSimpleOpt object
	CSimpleOpt csoArgs(iArgC, pwszArgV, g_Options);

	//Run the main program logic
	iReturn = !Run(csoArgs);

	//Free the command line
	if (pwszArgV) { LocalFree(pwszArgV); pwszArgV = NULL; }
	//Terminate WTL module
	_Module.Term();
	//Unload COM
	CoUninitialize();

	//Let the caller know the results
	if (iReturn == 0) {wprintf(L"Success!"); }
	return iReturn;
}