#include "..\inc\ez1OUProvider.h"
#include "..\inc\MergeWithMAPISVC.h"

// macro to get allocated string lengths
#define CCH(string) (sizeof((string))/sizeof(TCHAR))

// Note: when the user loads MAPI manually, hModMSMAPI and hModMAPI will be the same
HMODULE	hModMSMAPI = NULL; // Address of Outlook's MAPI
HMODULE	hModMAPI = NULL; // Address of MAPI32 in System32
HMODULE	hModMAPIStub = NULL; // Address of the MAPI stub library

// All of these get loaded from a MAPI DLL:
LPFGETCOMPONENTPATH			pfnFGetComponentPath = NULL;

STDMETHODIMP MergeWithMAPISVC();
LPTSTR GenerateProviderPath();
STDMETHODIMP HrSetProfileParameters(SERVICESINIREC *lpServicesIni);
void GetMAPISVCPath(LPTSTR szMAPIDir, ULONG cchMAPIDir);
void GetMAPIPath(LPTSTR szClient, LPTSTR szMAPIPath, ULONG cchMAPIPath);
void GetMapiMsiIds(LPTSTR szClient, LPTSTR* lpszComponentID, LPTSTR* lpszAppLCID, LPTSTR* lpszOfficeLCID);
HKEY GetMailKey(LPTSTR szClient);
LONG GetRegistryValue( IN HKEY hKey,  IN LPCTSTR lpszValue, OUT DWORD* lpType, OUT LPVOID* lppData);
BOOL GetComponentPath(LPTSTR szComponent, LPTSTR szQualifier, TCHAR* szDllPath, DWORD cchDLLPath);
void LoadGetComponentPath();
HMODULE LoadFromSystemDir(LPTSTR szDLLName);
HMODULE MyLoadLibrary(LPCTSTR lpLibFileName);


//"c:\Windows\SysWOW64\rundll32.exe" "c:\tempQC\Qualcomm\IPDocs\Debug\IPDOCS32.dll" MergeWithMAPISVC
STDMETHODIMP MergeWithMAPISVC()
{
	Log(true, "MergeWithMAPISVC\n");
	Log(true, "DLL install path is %s\n", aWrapPSTServicesIni[2].lpszValue);
	LPTSTR szPath = GenerateProviderPath();
	Log(true, "Better DLL install path is %s\n", szPath);
	aWrapPSTServicesIni[2].lpszValue = szPath;
	aWrapPSTServicesIni[9].lpszValue = szPath;
	
	HRESULT hRes = HrSetProfileParameters(aWrapPSTServicesIni);
	
	delete[] szPath;
	return S_OK;
}

//"c:\Windows\SysWOW64\rundll32.exe" "c:\tempQC\Qualcomm\IPDocs\Debug\IPDOCS32.dll" RemoveFromMAPISVC
STDMETHODIMP RemoveFromMAPISVC()
{
	Log(true, "RemoveFromMAPISVC removing IPDOCS\n");
	return HrSetProfileParameters(aREMOVE_WrapPSTServicesIni);
}


LPTSTR GenerateProviderPath()
{
	TCHAR szFilePath[MAX_PATH] = { 0 };
	DWORD dwDir = NULL;
	HMODULE hMod = NULL;
	hMod = GetModuleHandle("IPDOCS32.dll");
	if (hMod)
	{
		dwDir = GetModuleFileName(hMod, szFilePath, _countof(szFilePath));
		TCHAR* szExt = _tcsstr(szFilePath, "32.dll");
		if (szExt)
		{
			StringCchPrintf(szExt, 6, ".dll");
			size_t cchStr = 0;
			StringCchLength(szFilePath, MAX_PATH, &cchStr);
			LPTSTR szPath = new TCHAR[cchStr + 1];
			if (szPath)
			{
				StringCchCopy(szPath, cchStr + 1, szFilePath);
				return szPath;
			}
		}
	}

	return NULL;
}

// $--HrSetProfileParameters----------------------------------------------
// Add values to MAPISVC.INF
// -----------------------------------------------------------------------------
STDMETHODIMP HrSetProfileParameters(SERVICESINIREC *lpServicesIni)
{
	HRESULT	hRes = S_OK;
	TCHAR	szSystemDir[MAX_PATH + 1] = { 0 };
	TCHAR	szServicesIni[MAX_PATH + 12] = { 0 }; // 12 = space for "MAPISVC.INF"
	UINT	n = 0;
	TCHAR	szPropNum[10] = { 0 };

	Log(true, _T("HrSetProfileParameters()\n"));

	if (!lpServicesIni) return MAPI_E_INVALID_PARAMETER;

	GetMAPISVCPath(szServicesIni, CCH(szServicesIni));

	if (!szServicesIni[0])
	{
		UINT uiLen = 0;
		uiLen = GetSystemDirectory(szSystemDir, CCH(szSystemDir));
		if (!uiLen)
			return MAPI_E_CALL_FAILED;

		Log(true, _T("Writing to this directory: \"%s\"\n"), szSystemDir);

		hRes = StringCchPrintf(
			szServicesIni,
			CCH(szServicesIni),
			_T("%s\\%s"),
			szSystemDir,
			_T("MAPISVC.INF"));
	}

	Log(true, _T("Writing to this file: \"%s\"\n"), szServicesIni);

	//
	// Loop through and add items to MAPISVC.INF
	//

	n = 0;

	while (lpServicesIni[n].lpszSection != NULL)
	{
		LPTSTR lpszProp = lpServicesIni[n].lpszKey;
		LPTSTR lpszValue = lpServicesIni[n].lpszValue;

		// Switch the property if necessary

		if ((lpszProp == NULL) && (lpServicesIni[n].ulKey != 0))
		{

			hRes = StringCchPrintf(
				szPropNum,
				CCH(szPropNum),
				_T("%lx"),
				lpServicesIni[n].ulKey);

			if (SUCCEEDED(hRes))
				lpszProp = szPropNum;
		}

		//
		// Write the item to MAPISVC.INF
		//

		WritePrivateProfileString(
			lpServicesIni[n].lpszSection,
			lpszProp,
			lpszValue,
			szServicesIni);
		n++;
	}

	// Flush the information - ignore the return code
	WritePrivateProfileString(NULL, NULL, NULL, szServicesIni);

	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
// Function name   : GetMAPISVCPath
// Description     : This will get the correct path to the MAPISVC.INF file.
// Return type     : void
// Argument        : LPSTR szMAPIDir - Buffer to hold the path to the MAPISVC file.
//                   ULONG cchMAPIDir - size of the buffer
void GetMAPISVCPath(LPTSTR szMAPIDir, ULONG cchMAPIDir)
{
	Log(true, "Enter GetMAPISVCPath\n");

	GetMAPIPath("Microsoft Outlook", szMAPIDir, cchMAPIDir);

	// We got the path to msmapi32.dll - need to strip it
	if (szMAPIDir[0] != _T('\0'))
	{
		LPTSTR lpszSlash = NULL;
		LPTSTR lpszCur = szMAPIDir;

		for (lpszSlash = lpszCur; *lpszCur; lpszCur = lpszCur++)
		{
			if (*lpszCur == _T('\\')) lpszSlash = lpszCur;
		}
		*lpszSlash = _T('\0');
	}

	if (szMAPIDir[0] == _T('\0'))
	{
		Log(true, _T("FGetComponentPath failed, loading system directory\n"));
		// Fall back on System32
		UINT uiLen = 0;
		uiLen = GetSystemDirectory(szMAPIDir, cchMAPIDir);
	}

	if (szMAPIDir[0] != _T('\0'))
	{
		Log(true, _T("Using directory: %s\n"), szMAPIDir);
		StringCchPrintf(
			szMAPIDir,
			cchMAPIDir,
			_T("%s\\%s"),
			szMAPIDir,
			_T("MAPISVC.INF"));
	}
}
void GetMAPIPath(LPTSTR szClient, LPTSTR szMAPIPath, ULONG cchMAPIPath)
{
	BOOL bRet = false;
	szMAPIPath[0] = '\0'; // Terminate String at pos 0 (safer if we fail below)
						  // Find some strings:
	LPTSTR szComponentID = NULL;
	LPTSTR szAppLCID = NULL;
	LPTSTR szOfficeLCID = NULL;

	GetMapiMsiIds(szClient, &szComponentID, &szAppLCID, &szOfficeLCID);
	
	if (szComponentID)
	{
		if (szAppLCID)
		{
			bRet = GetComponentPath(szComponentID, szAppLCID, szMAPIPath,cchMAPIPath);
		}
		if ((!bRet || szMAPIPath[0] == _T('\0')) && szOfficeLCID)
		{
			bRet = GetComponentPath(szComponentID,szOfficeLCID,szMAPIPath,cchMAPIPath);
		}
		if (!bRet || szMAPIPath[0] == _T('\0'))
		{
			bRet = GetComponentPath(szComponentID,NULL,szMAPIPath,cchMAPIPath);
		}
	}	

	delete[] szComponentID;
	delete[] szOfficeLCID;
	delete[] szAppLCID;
} // GetMAPIPath
  // Gets MSI IDs for the specified MAPI client, such as 'Microsoft Outlook' or 'ExchangeMAPI'
  // Pass NULL to get the IDs for the default MAPI client
  // Allocates with new, delete with delete[]
void GetMapiMsiIds(LPTSTR szClient, LPTSTR* lpszComponentID, LPTSTR* lpszAppLCID, LPTSTR* lpszOfficeLCID)
{
	Log(true, _T("GetMapiMsiIds()\n"));
	LONG lRes = S_OK;
	HKEY hKey = GetMailKey(szClient);

	if (hKey)
	{
		DWORD dwKeyType = NULL;

		if (lpszComponentID)
		{
			lRes = GetRegistryValue( hKey, _T("MSIComponentID"), &dwKeyType,(LPVOID*)lpszComponentID);
			Log(true, _T("MSIComponentID = %s\n"), *lpszComponentID ? *lpszComponentID : _T("<not found>"));
		}

		if (lpszAppLCID)
		{
			lRes = GetRegistryValue( hKey, _T("MSIApplicationLCID"), &dwKeyType, (LPVOID*)lpszAppLCID);
			Log(true, _T("MSIApplicationLCID = %s\n"), *lpszAppLCID ? *lpszAppLCID : _T("<not found>"));
		}

		if (lpszOfficeLCID)
		{
			lRes = GetRegistryValue(hKey,_T("MSIOfficeLCID"),&dwKeyType, (LPVOID*)lpszOfficeLCID);
			Log(true, _T("MSIOfficeLCID = %s\n"), *lpszOfficeLCID ? *lpszOfficeLCID : _T("<not found>"));
		}

		RegCloseKey(hKey);
	}
} // GetMapiMsiIds

  // Opens the mail key for the specified MAPI client, such as 'Microsoft Outlook' or 'ExchangeMAPI'
  // Pass NULL to open the mail key for the default MAPI client
HKEY GetMailKey(LPTSTR szClient)
{
	Log(true, _T("Enter GetMailKey(%s)\n"), szClient ? szClient : _T("Default"));
	HRESULT hRes = S_OK;
	LONG lRet = S_OK;
	HKEY hMailKey = NULL;
	BOOL bClientIsDefault = false;

	// If szClient is NULL, we need to read the name of the default MAPI client
	if (!szClient)
	{
		HKEY hDefaultMailKey = NULL;
		lRet = RegOpenKeyEx( HKEY_LOCAL_MACHINE, _T("Software\\Clients\\Mail"), NULL, KEY_READ, &hDefaultMailKey);
		if (hDefaultMailKey)
		{
			DWORD dwKeyType = NULL;
			lRet = GetRegistryValue( hDefaultMailKey, _T(""), &dwKeyType, (LPVOID*)&szClient);

			Log(true, _T("Default MAPI = %s\n"), szClient ? szClient : _T("Default"));
			bClientIsDefault = true;
			RegCloseKey(hDefaultMailKey);
		}
	}

	if (szClient)
	{
		TCHAR szMailKey[256];
		hRes = StringCchPrintf(szMailKey, CCH(szMailKey), _T("Software\\Clients\\Mail\\%s"), szClient);

		if (SUCCEEDED(hRes))
		{
			lRet = RegOpenKeyEx( HKEY_LOCAL_MACHINE, szMailKey, NULL, KEY_READ, &hMailKey);
		}
	}
	if (bClientIsDefault) delete[] szClient;

	return hMailKey;
} // GetMailKey


  // $--GetRegistryValue---------------------------------------------------------
  // Get a registry value - allocating memory using new to hold it.
  // -----------------------------------------------------------------------------
LONG GetRegistryValue(
	IN HKEY hKey, // the key.
	IN LPCTSTR lpszValue, // value name in key.
	OUT DWORD* lpType, // where to put type info.
	OUT LPVOID* lppData) // where to put the data.
{
	LONG lRes = ERROR_SUCCESS;

	Log(true, _T("GetRegistryValue(%s)\n"), lpszValue);

	*lppData = NULL;
	DWORD cb = NULL;

	// Get its size
	lRes = RegQueryValueEx(
		hKey,
		lpszValue,
		NULL,
		lpType,
		NULL,
		&cb);

	// only handle types we know about - all others are bad
	if (ERROR_SUCCESS == lRes && cb &&
		(REG_SZ == *lpType || REG_DWORD == *lpType || REG_MULTI_SZ == *lpType))
	{
		*lppData = new BYTE[cb];

		if (*lppData)
		{
			// Get the current value
			lRes = RegQueryValueEx(
				hKey,
				lpszValue,
				NULL,
				lpType,
				(unsigned char*)*lppData,
				&cb);

			if (ERROR_SUCCESS != lRes)
			{
				delete[] * lppData;
				*lppData = NULL;
			}
		}
	}
	else lRes = ERROR_INVALID_DATA;

	return lRes;
}

BOOL GetComponentPath( LPTSTR szComponent, LPTSTR szQualifier, TCHAR* szDllPath, DWORD cchDLLPath)
{
	BOOL bRet = false;
	LoadGetComponentPath();

	if (!pfnFGetComponentPath) return false;
#ifdef UNICODE
	int iRet = 0;
	CHAR szAsciiPath[MAX_PATH] = { 0 };
	char szAnsiComponent[MAX_PATH] = { 0 };
	char szAnsiQualifier[MAX_PATH] = { 0 };
	iRet = WideCharToMultiByte(CP_ACP, 0, szComponent, (int)-1, szAnsiComponent, CCH(szAnsiComponent), NULL, NULL);
	iRet = WideCharToMultiByte(CP_ACP, 0, szQualifier, (int)-1, szAnsiQualifier, CCH(szAnsiQualifier), NULL, NULL);

	bRet = pfnFGetComponentPath(
		szAnsiComponent,
		szAnsiQualifier,
		szAsciiPath,
		CCH(szAsciiPath),
		TRUE);

	iRet = MultiByteToWideChar(
		CP_ACP,
		0,
		szAsciiPath,
		CCH(szAsciiPath),
		szDllPath,
		cchDLLPath);
#else
	bRet = pfnFGetComponentPath(
		szComponent,
		szQualifier,
		szDllPath,
		cchDLLPath,
		TRUE);
#endif
	return bRet;
} // GetComponentPath

void LoadGetComponentPath()
{
	HRESULT hRes = S_OK;

	if (pfnFGetComponentPath) return;

	if (!hModMAPI) hModMAPI = LoadFromSystemDir(_T("mapi32.dll"));
	if (hModMAPI)
	{
		pfnFGetComponentPath = (LPFGETCOMPONENTPATH)GetProcAddress(
			hModMAPI,
			"FGetComponentPath");
		hRes = S_OK;
	}
	if (!pfnFGetComponentPath)
	{
		if (!hModMAPIStub) hModMAPIStub = LoadFromSystemDir(_T("mapistub.dll"));
		if (hModMAPIStub)
		{
			pfnFGetComponentPath = (LPFGETCOMPONENTPATH)GetProcAddress(
				hModMAPIStub,
				"FGetComponentPath");
		}
	}

	Log(true, _T("FGetComponentPath loaded at 0x%08X\n"), pfnFGetComponentPath);
} // LoadGetComponentPath

HMODULE LoadFromSystemDir(LPTSTR szDLLName)
{
	if (!szDLLName) return NULL;

	HRESULT	hRes = S_OK;
	HMODULE	hModRet = NULL;
	TCHAR	szDLLPath[MAX_PATH] = { 0 };
	UINT	uiRet = NULL;

	static TCHAR	szSystemDir[MAX_PATH] = { 0 };
	static BOOL		bSystemDirLoaded = false;

	Log(true, _T("LoadFromSystemDir - loading \"%s\"\n"), szDLLName);

	if (!bSystemDirLoaded)
	{
		uiRet = GetSystemDirectory(szSystemDir, MAX_PATH);
		bSystemDirLoaded = true;
	}

	hRes = StringCchPrintf(szDLLPath, CCH(szDLLPath), _T("%s\\%s"), szSystemDir, szDLLName);
	Log(true, _T("LoadFromSystemDir - loading from \"%s\"\n"), szDLLPath);
	hModRet = MyLoadLibrary(szDLLPath);

	return hModRet;
}

// Exists to allow some logging
HMODULE MyLoadLibrary(LPCTSTR lpLibFileName)
{
	HMODULE hMod = NULL;
	Log(true, _T("MyLoadLibrary - loading \"%s\"\n"), lpLibFileName);
	hMod = LoadLibrary(lpLibFileName);
	if (hMod)
	{
		Log(true, _T("MyLoadLibrary - \"%s\" loaded at 0x%08X\n"), lpLibFileName, hMod);
	}
	else
	{
		Log(true, _T("MyLoadLibrary - \"%s\" failed to load\n"), lpLibFileName);
	}
	return hMod;
}
