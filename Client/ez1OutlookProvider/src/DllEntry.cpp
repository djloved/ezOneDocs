#include "..\inc\ez1OUProvider.h"

#include <iostream>
using namespace std;
using namespace ez1OutlookProvider;

BOOLEAN WINAPI DllMain(IN HINSTANCE hDllHandle,
	IN DWORD     nReason,
	IN LPVOID    Reserved)
{	
	BOOLEAN bSuccess = TRUE;
	//  Perform global initialization.
	switch (nReason)
	{
	case DLL_PROCESS_ATTACH:
		//  For optimization.
		Log(true, "dll attached \n");
		//		DisableThreadLibraryCalls(hDllHandle);
		break;
	case DLL_PROCESS_DETACH:
		Log(true, "dll deattached \n");
		break;
	}
	return true;

}
//  end DllMain
//
LPALLOCATEBUFFER g_lpAllocateBuffer = NULL;
LPALLOCATEMORE g_lpAllocateMore = NULL;
LPFREEBUFFER g_lpFreeBuffer = NULL;

STDINITMETHODIMP test1()
{
	Log(true, "MSProviderInit");
	cout << "MSProviderInit";
	return S_OK;
}
STDINITMETHODIMP MSProviderInit(
	HINSTANCE				hInstance,
	LPMALLOC				lpMalloc,
	LPALLOCATEBUFFER		lpAllocateBuffer,
	LPALLOCATEMORE			lpAllocateMore,
	LPFREEBUFFER			lpFreeBuffer,
	ULONG					ulFlags,
	ULONG					ulMAPIVer,
	ULONG FAR *			lpulProviderVer,
	LPMSPROVIDER FAR *		lppMSProvider)
{
	Log(true, "MSProviderInit function called\n"); 
	if (!lppMSProvider || !lpulProviderVer) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	*lppMSProvider = NULL;
	*lpulProviderVer = CURRENT_SPI_VERSION;
	if (ulMAPIVer < CURRENT_SPI_VERSION)
	{
		Log(true, "MSProviderInit: The version of the subsystem cannot handle this version of the provider\n");
		return MAPI_E_VERSION;
	}

	Log(true, "MSProviderInit: saving off memory management routines\n");
	g_lpAllocateBuffer = lpAllocateBuffer;
	g_lpAllocateMore = lpAllocateMore;
		g_lpFreeBuffer = lpFreeBuffer;


	ez1MSProvider *lpMSProvider = NULL;
	hRes = ez1MSProvider::Create(ulFlags, &lpMSProvider);
	return lpMSProvider->QueryInterface(IID_IMSProvider, (void**)lppMSProvider);

}

HRESULT STDAPICALLTYPE ServiceEntry(
	HINSTANCE hInstance,
	LPMALLOC lpMalloc,
	LPMAPISUP lpMAPISup,
	ULONG_PTR ulUIParam,
	ULONG ulFlags,
	ULONG ulContext,
	ULONG cValues,
	LPSPropValue lpProps,
	LPPROVIDERADMIN lpProviderAdmin,
	LPMAPIERROR FAR * lppMapiError
)
{
	HRESULT hRes = S_OK;
	Log(true, "ServiceEntry\n");

	return hRes;
}