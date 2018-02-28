#pragma once
#ifndef __QCSLogon__
#define __QCSLogon__

#include "ez1OUProvider.h"
namespace ez1OutlookProvider
{
	class ez1Logon : public IMSLogon
	{
	public:
		STDMETHODIMP QueryInterface(REFIID  riid, LPVOID *ppvObj);
		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();
		MAPI_IMSLOGON_METHODS(IMPL);
	public:
		ez1Logon();
		~ez1Logon();
	private:
		ULONG              m_cRef;
	};
}


#endif // !MERGE_WITH_MAPI_SVC

