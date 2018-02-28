#pragma once
#ifndef __QCSMsgStore_H__
#define __QCSMsgStore_H__

#include "ez1OUProvider.h"
namespace ez1OutlookProvider
{
	//https://msdn.microsoft.com/en-us/library/office/cc815283.aspx
	//Provides access to message store information and to messages and folders.
	class ez1MsgStore :public IMsgStore
	{
	public:
		ez1MsgStore();
		~ez1MsgStore();
	public:
		///////////////////////////////////////////////////////////////////////////////
		// Interface virtual member functions
		//
		STDMETHODIMP QueryInterface(REFIID riid, LPVOID *ppvObj);
		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();
		MAPI_IMAPIPROP_METHODS(IMPL);
		MAPI_IMSGSTORE_METHODS(IMPL);
	private:
		ULONG              m_cRef;
	};
}

#endif QCS_MSG_STORE