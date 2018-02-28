#include "..\inc\ez1OUProvider.h"
#include <MAPIUtil.h>

namespace ez1OutlookProvider
{
	/*HRESULT MapiUtil::HrGetOneProp(LPMAPIPROP lpMapiProp,
		ULONG ulPropTag,
		LPSPropValue FAR * lppProp)
	{
		CHECKLOAD(pfnHrGetOneProp);
		if (pfnHrGetOneProp) return pfnHrGetOneProp(lpMapiProp, ulPropTag, lppProp);
		return MAPI_E_CALL_FAILED;
	}*/


	//HRESULT QCSMapiUtil::GetGlobalProfileProperties(LPMAPISUP lpMAPISup, struct sGlobalProfileProps* lpsProfileProps)
	//{
	//	HRESULT			hr = hrSuccess;
	//	LPPROFSECT		lpGlobalProfSect = NULL;

	//	// open profile
	//	hr = lpMAPISup->OpenProfileSection((LPMAPIUID)pbGlobalProfileSectionGuid, MAPI_MODIFY, &lpGlobalProfSect);
	//	if (hr != hrSuccess)
	//		goto exit;

	//	hr = QCSMapiUtil::GetGlobalProfileProperties(lpGlobalProfSect, lpsProfileProps);
	//	if (hr != hrSuccess)
	//		goto exit;

	//exit:
	//	if (lpGlobalProfSect)
	//		lpGlobalProfSect->Release();

	//	return hr;
	//}

	//HRESULT QCSMapiUtil::GetGlobalProfileProperties(LPPROFSECT lpGlobalProfSect, struct sGlobalProfileProps* lpsProfileProps)
	//{
	//	HRESULT			hr = hrSuccess;
	//	LPSPropValue	lpsPropArray = NULL;
	//	ULONG			cValues = 0;
	//	LPSPropValue	lpsEMSPropArray = NULL;
	//	LPSPropValue	lpPropEMS = NULL;
	//	ULONG			cEMSValues = 0;
	//	LPSPropValue	lpProp = NULL;
	//	bool			bIsEMS = false;

	//	if (lpGlobalProfSect == NULL || lpsProfileProps == NULL)
	//	{
	//		hr = MAPI_E_INVALID_OBJECT;
	//		goto exit;
	//	}

	//	//if (HrGetOneProp(lpGlobalProfSect, PR_PROFILE_UNRESOLVED_NAME, &lpPropEMS) == hrSuccess || g_ulLoadsim) {
	//	if (HrGetOneProp(lpGlobalProfSect, PR_PROFILE_UNRESOLVED_NAME, &lpPropEMS) == hrSuccess ) {
	//		bIsEMS = true;
	//	}

	//	if (bIsEMS) {
	//		SizedSPropTagArray(4, sptaEMSProfile) = { 4,{ PR_PROFILE_NAME_A, PR_PROFILE_UNRESOLVED_SERVER, PR_PROFILE_UNRESOLVED_NAME, PR_PROFILE_USER } };

	//		// This is an emulated MSEMS store. Get the properties we need and convert them to ZARAFA-style properties
	//		hr = lpGlobalProfSect->GetProps((LPSPropTagArray)&sptaEMSProfile, 0, &cEMSValues, &lpsEMSPropArray);
	//		if (FAILED(hr))
	//			goto exit;

	//		//hr = ConvertMSEMSProps(cEMSValues, lpsEMSPropArray, &cValues, &lpsPropArray);
	//		//if (FAILED(hr))
	//		//	goto exit;
	//	}
	//	else {
	//		// Get the properties we need directly from the global profile section
	//	/*	hr = lpGlobalProfSect->GetProps((LPSPropTagArray)&sptaZarafaProfile, 0, &cValues, &lpsPropArray);
	//		if (FAILED(hr))
	//			goto exit;*/
	//	}

	//	/*
	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_PATH)) != NULL)
	//		lpsProfileProps->strServerPath = lpProp->Value.lpszA;

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_PROFILE_NAME_A)) != NULL)
	//		lpsProfileProps->strProfileName = lpProp->Value.lpszA;

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_USERNAME_W)) != NULL)
	//		lpsProfileProps->strUserName = convstring::from_SPropValue(lpProp);
	//	else if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_USERNAME_A)) != NULL)
	//		lpsProfileProps->strUserName = convstring::from_SPropValue(lpProp);

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_USERPASSWORD_W)) != NULL)
	//		lpsProfileProps->strPassword = convstring::from_SPropValue(lpProp);
	//	else if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_USERPASSWORD_A)) != NULL)
	//		lpsProfileProps->strPassword = convstring::from_SPropValue(lpProp);

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_IMPERSONATEUSER_W)) != NULL)
	//		lpsProfileProps->strImpersonateUser = convstring::from_SPropValue(lpProp);
	//	else if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_IMPERSONATEUSER_A)) != NULL)
	//		lpsProfileProps->strImpersonateUser = convstring::from_SPropValue(lpProp);

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_FLAGS)) != NULL)
	//		lpsProfileProps->ulProfileFlags = lpProp->Value.ul;
	//	else
	//		lpsProfileProps->ulProfileFlags = 0;

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_SSLKEY_FILE)) != NULL)
	//		lpsProfileProps->strSSLKeyFile = lpProp->Value.lpszA;

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_SSLKEY_PASS)) != NULL)
	//		lpsProfileProps->strSSLKeyPass = lpProp->Value.lpszA;

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_PROXY_HOST)) != NULL)
	//		lpsProfileProps->strProxyHost = lpProp->Value.lpszA;

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_PROXY_PORT)) != NULL)
	//		lpsProfileProps->ulProxyPort = lpProp->Value.ul;
	//	else
	//		lpsProfileProps->ulProxyPort = 0;

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_PROXY_FLAGS)) != NULL)
	//		lpsProfileProps->ulProxyFlags = lpProp->Value.ul;
	//	else
	//		lpsProfileProps->ulProxyFlags = 0;

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_PROXY_USERNAME)) != NULL)
	//		lpsProfileProps->strProxyUserName = lpProp->Value.lpszA;

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_PROXY_PASSWORD)) != NULL)
	//		lpsProfileProps->strProxyPassword = lpProp->Value.lpszA;

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_CONNECTION_TIMEOUT)) != NULL)
	//		lpsProfileProps->ulConnectionTimeOut = lpProp->Value.ul;
	//	else
	//		lpsProfileProps->ulConnectionTimeOut = 10;

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_OFFLINE_PATH_W)) != NULL)
	//		lpsProfileProps->strOfflinePath = convstring::from_SPropValue(lpProp);
	//	else if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_OFFLINE_PATH_A)) != NULL)
	//		lpsProfileProps->strOfflinePath = convstring::from_SPropValue(lpProp);

	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_STATS_SESSION_CLIENT_APPLICATION_VERSION)) != NULL)
	//		lpsProfileProps->strClientAppVersion = lpProp->Value.lpszA;
	//	if ((lpProp = PpropFindProp(lpsPropArray, cValues, PR_EC_STATS_SESSION_CLIENT_APPLICATION_MISC)) != NULL)
	//		lpsProfileProps->strClientAppMisc = lpProp->Value.lpszA;

	//	lpsProfileProps->bIsEMS = bIsEMS;
	//	*/
	//	hr = hrSuccess;

	//exit:
	//	/*if (lpPropEMS)
	//		MAPIFreeBuffer(lpPropEMS);

	//	if (lpsPropArray)
	//		MAPIFreeBuffer(lpsPropArray);

	//	if (lpsEMSPropArray)
	//		MAPIFreeBuffer(lpsEMSPropArray);*/

	//	return hr;
	//}

}