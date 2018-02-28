#pragma once
#ifndef __ez1MSProvider_H__
#define __ez1MSProvider_H__

#include "ez1OUProvider.h"

/*
https://msdn.microsoft.com/en-us/library/office/cc765644.aspx
Provides access to a message store provider through a message store provider object. 
This message store provider object is returned at provider logon by the message store provider's MSProviderInit entry point function.
The message store provider object is primarily used by client applications and the MAPI spooler to open message stores.
*/
/*
Loading Message Store Providers
https://msdn.microsoft.com/en-us/library/cc839581(v=office.15).aspx

1.The client calls IMAPISession::OpenMsgStore.
2.If the message store is not already open, MAPI loads the store provider's DLL and calls the DLL's MSProviderInit entry point. If the message store is already open, MAPI skips steps 2 and 3, and then uses the existing IMSProvider : IUnknown interface to complete step 4.
3.MSProviderInit creates and returns an IMSProvider object.
4.MAPI calls IMSProvider::Logon, passing the client application's message store entry identifier.
5.IMSProvider::Logon creates and returns an IMSLogon : IUnknown interface and an IMsgStore : IMAPIProp interface, and then calls the IUnknown::AddRef method on its IMAPISupport : IUnknown interface. If the client's message store entry identifier refers to a message store that is already open, the message store provider can return existing IMSLogon and IMsgStore interfaces and does not need to call AddRef on its support object
6.If the client did not set the MAPI_NO_MAIL flag when it logged on and it did not set the MDB_NO_MAIL in step 1, MAPI gives the message store's entry identifier to the MAPI spooler so the MAPI spooler can log on to the message store.
7.MAPI returns the IMsgStore interface to the client.
8.The MAPI spooler calls IMSProvider::SpoolerLogon.
9.IMSProvider::SpoolerLogon returns the same IMSLogon and IMsgStore interfaces from step 5.
*/
namespace ez1OutlookProvider
{
	class ez1MSProvider : public IMSProvider
	{
	public:
		///////////////////////////////////////////////////////////////////////////////
		// Interface virtual member functions
		//
		/*STDMETHODIMP QueryInterface (REFIID  riid, LPVOID *ppvObj);
		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();
		MAPI_IMSPROVIDER_METHODS(IMPL);*/
		virtual STDMETHODIMP_(ULONG)	AddRef(void);
		virtual STDMETHODIMP_(ULONG)	Release(void) ;
		virtual STDMETHODIMP			QueryInterface(REFIID refiid, void **lppInterface);

		// MSProvider
		virtual STDMETHODIMP Shutdown(ULONG * lpulFlags);
		virtual STDMETHODIMP Logon(LPMAPISUP lpMAPISup, ULONG ulUIParam, LPTSTR lpszProfileName, ULONG cbEntryID, LPENTRYID lpEntryID, ULONG ulFlags, LPCIID lpInterface, ULONG *lpcbSpoolSecurity, LPBYTE *lppbSpoolSecurity, LPMAPIERROR *lppMAPIError, LPMSLOGON *lppMSLogon, LPMDB *lppMDB);
		virtual STDMETHODIMP SpoolerLogon(LPMAPISUP lpMAPISup, ULONG ulUIParam, LPTSTR lpszProfileName, ULONG cbEntryID, LPENTRYID lpEntryID, ULONG ulFlags, LPCIID lpInterface, ULONG lpcbSpoolSecurity, LPBYTE lppbSpoolSecurity, LPMAPIERROR *lppMAPIError, LPMSLOGON *lppMSLogon, LPMDB *lppMDB);
		virtual STDMETHODIMP CompareStoreIDs(ULONG cbEntryID1, LPENTRYID lpEntryID1, ULONG cbEntryID2, LPENTRYID lpEntryID2, ULONG ulFlags, ULONG *lpulResult);

		static HRESULT Create(ULONG ulFlags, ez1MSProvider **lppMSProvider);

	public:
		ez1MSProvider(ULONG ulFlags);
		~ez1MSProvider();
	private:
		ULONG              m_cRef;
	protected:
		ULONG			m_ulFlags;
	};
}

#endif