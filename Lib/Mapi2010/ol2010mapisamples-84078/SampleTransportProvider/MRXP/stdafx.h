// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

// Modify the following defines if you have to target a version prior to the ones specified below.
// Refer to MSDN for the latest information on corresponding values for different versions.
#ifndef WINVER				// Allow use of features specific to Windows XP or later.
#define WINVER 0x0501		// Change this to the appropriate value to target other versions of Windows.
#endif

#ifndef _WIN32_WINNT		// Allow use of features specific to Windows XP or later.

#define _WIN32_WINNT 0x0501	// Change this to the appropriate value to target other versions of Windows.
#endif


#ifndef _WIN32_WINDOWS		// Allow use of features specific to Windows 98 or later.
#define _WIN32_WINDOWS 0x0410 // Change this to the appropriate value to target Windows Me or later.
#endif

#ifndef _WIN32_IE			// Allow use of features specific to Internet Explorer 6.0 or later.
#define _WIN32_IE 0x0600	// Change this to the appropriate value to target other versions of Internet Explorer.
#endif

#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <WINDOWS.H>
#include <MAPIWIN.H>
#include <MAPISPI.H>
#include <MAPIUTIL.H>
#include <MAPIVAL.H>
#include <WINDOWSX.H>

// TODO: reference additional headers your program requires here
#include "mrxp.h"
#include "Output.h"
#include "ImportProcs.h"