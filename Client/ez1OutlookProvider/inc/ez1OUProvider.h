// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once
#ifndef __ez1OUProvider__H__
#define __ez1OUProvider__H__

//#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:

#include <stdio.h>
#include <WINDOWS.H>
#include <COMMCTRL.H>
#include <tchar.h>
#include <strsafe.h>
#include <exception>
#include <string>

#include <MAPISPI.H>
#include <MAPIDefS.h>
#include <MAPIUtil.h>

#define INITGUID

#define USES_IID_IMAPIProp
#define USES_IID_IMAPIPropData
#define USES_IID_IMSProvider
#define USES_IID_IMSLogon
#define USES_IID_IMsgStore
#define USES_IID_IMAPIContainer
#define USES_IID_IMAPIFolder
#define USES_IID_IMessage
#define USES_IID_IAttachment
#define USES_IID_IMAPITable
#define USES_IID_IMAPITableData
#define USES_IID_IMAPITable
#define USES_IID_IMAPIStatus
#define USES_IID_IMAPIControl
#define USES_IID_IMAPIForm
#define USES_IID_IMAPIFormAdviseSink
#define USES_IID_IPersistMessage
#define USES_IID_IMAPISup
#define USES_IID_IMAPISession
#define USES_IID_IMAPIFormAdviseSink

#include <initguid.h>
#include <mapiguid.h>
#include "EdkMdb.h"
#include <MAPIUtil.h>
#include "MergeWithMAPISVC.h"
#include "ez1MSProvider.h"
#include "ez1MsgStore.h"
#include "ez1Logon.h"
#include "ez1MapiUtil.h"

using namespace ez1OutlookProvider;

namespace ez1OutlookProvider
{
	void DeInitLogging();
	void Log(BOOL bPrintThreadTime, LPCTSTR szMsg, ...);
	void LogREFIID(REFIID riid);
	
	class QCSProvider;
	typedef MSPROVIDERINIT FAR *LPMSPROVIDERINIT;
	#define MDB_OST_LOGON_UNICODE	((ULONG) 0x00000800)
	#define MDB_OST_LOGON_ANSI		((ULONG) 0x00001000)

	class QCSMsgStore;	
	class QCSLogon;
	class QCSMapiUtil;
}

#endif


