#pragma once
#ifndef __MapiUtil_H__
#define __MapiUtil_H__

#include "ez1OUProvider.h"
#include <MAPIUtil.h>

/**
* Return interface pointer on a specific interface query.
* @param[in]	_guid	The interface guid.
* @param[in]	_interface	The class which implements the interface
* @note guid variable must be named 'refiid', return variable must be named lppInterface.
*/
#define REGISTER_INTERFACE(_guid, _interface)	\
	if (refiid == (_guid)) {				 	\
		AddRef();								\
		*lppInterface = (void*)(_interface);	\
		return hrSuccess;						\
	}

/**
* Return interface pointer on a specific interface query without incrementing the refcount.
* @param[in]	_guid	The interface guid.
* @param[in]	_interface	The class which implements the interface
* @note guid variable must be named 'refiid', return variable must be named lppInterface.
*/
#define REGISTER_INTERFACE_NOREF(_guid, _interface)	\
	if (refiid == (_guid)) {				 		\
		AddRef();									\
		*lppInterface = (void*)(_interface);		\
		return hrSuccess;							\
	}


namespace ez1OutlookProvider
{
	struct sGlobalProfileProps
	{
		std::string		strServerPath;
		std::string		strProfileName;
		std::wstring		strUserName;
		std::wstring		strPassword;
		std::wstring    	strImpersonateUser;
		ULONG			ulProfileFlags;
		std::string		strSSLKeyFile;
		std::string		strSSLKeyPass;
		ULONG			ulConnectionTimeOut;
		ULONG			ulProxyFlags;
		std::string		strProxyHost;
		ULONG			ulProxyPort;
		std::string		strProxyUserName;
		std::string		strProxyPassword;
		std::string		strOfflinePath;
		bool			bIsEMS;
		std::string		strClientAppVersion;
		std::string		strClientAppMisc;
	};
	class ez1MapiUtil
	{
	public:		
		
		static HRESULT __stdcall GetGlobalProfileProperties(LPPROFSECT lpGlobalProfSect, struct sGlobalProfileProps* lpsProfileProps);
		static HRESULT __stdcall GetGlobalProfileProperties(LPMAPISUP lpMAPISup, struct sGlobalProfileProps* lpsProfileProps);

	};
}

#endif