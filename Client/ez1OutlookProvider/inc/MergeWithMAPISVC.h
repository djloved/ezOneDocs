#pragma once
#ifndef __MergeWithMAPISVC__
#define __MergeWithMAPISVC__

#include "ez1OUProvider.h"

typedef BOOL(STDAPICALLTYPE FGETCOMPONENTPATH)
(LPSTR szComponent,
	LPSTR szQualifier,
	LPSTR szDllPath,
	DWORD cchBufferSize,
	BOOL fInstall);
typedef FGETCOMPONENTPATH FAR * LPFGETCOMPONENTPATH;

typedef struct
{
	LPTSTR lpszSection;
	LPTSTR lpszKey;
	ULONG ulKey;
	LPTSTR lpszValue;
} SERVICESINIREC;

static SERVICESINIREC aWrapPSTServicesIni[] =
{
	///**/
	//{ _T("Services"),	_T("IPDOCS"), 0L, _T("QC IPIT DOCS") },

	//{ _T("IPDOCS"),	_T("PR_DISPLAY_NAME"),			0L, _T("QC IPIT DOCS") },
	//{ _T("IPDOCS"),	_T("PR_SERVICE_DLL_NAME"),		0L, _T("IPDOCS.DLL") },
	//{ _T("IPDOCS"),	_T("PR_SERVICE_ENTRY_NAME"),	0L, _T("ServiceEntry") },
	//{ _T("IPDOCS"),	_T("PR_RESOURCE_FLAGS"),		0L, _T("SERVICE_NO_PRIMARY_IDENTITY|SERVICE_SINGLE_COPY") },
	//{ _T("IPDOCS"),	_T("Providers"),				0L,	_T("MS_IPDOCS_P") },
	//{ _T("IPDOCS"),	_T("PR_SERVICE_SUPPORT_FILES"),	0L, _T("IPDOCS.DLL") },
	//{ _T("IPDOCS"),	_T("PR_SERVICE_DELETE_FILES"),	0L,	_T("IPDOCS.DLL") },

	//{ _T("MS_IPDOCS_P"),	_T("PR_RESOURCE_TYPE"),			0L, _T("MAPI_STORE_PROVIDER") },
	//{ _T("MS_IPDOCS_P"),	_T("PR_PROVIDER_DLL_NAME"),		0L, _T("IPDOCS.DLL") },
	//{ _T("MS_IPDOCS_P"),	_T("PR_RESOURCE_FLAGS"),		0L, _T("STATUS_NO_DEFAULT_STORE|SERVICE_SINGLE_COPY") },
	//{ _T("MS_IPDOCS_P"),	_T("PR_DISPLAY_NAME"),			0L, _T("QC IPIT DOCS") },
	//{ _T("MS_IPDOCS_P"),	_T("PR_PROVIDER_DISPLAY"),		0L, _T("QC IPIT DOCS Provider") },
	////	{_T("MS_WRAPPST_P"),	NULL,							0x67020003, _T("00000010")}, // uncomment to use Unicode PST

	{ NULL, NULL, 0L, NULL }
};

static SERVICESINIREC aREMOVE_WrapPSTServicesIni[] =
{
	/*{ _T("Services"),      _T("IPDOCS"), 0L, NULL },

	{ _T("IPDOCS"),       NULL,          0L, NULL },

	{ _T("MS_IPDOCS_P"),  NULL,          0L, NULL },*/

	{ NULL,                NULL,          0L, NULL }
};

LPTSTR GenerateProviderPath();
STDMETHODIMP MergeWithMAPISVC();
STDMETHODIMP RemoveFromMAPISVC();

#endif // !MERGE_WITH_MAPI_SVC
