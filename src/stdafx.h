// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#include <stdio.h>
#include <tchar.h>

#define _ATL_APARTMENT_THREADED

#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit

#include <vector>
#include <string>
#include <memory>
#include <utility>

#include <cassert>

// TODO: reference additional headers your program requires here
#include <atlbase.h>
#include <atlcom.h>

#include <windows.h>

#include "utils.h"