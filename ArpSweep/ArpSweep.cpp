/* Copyright (c) Constantine V. Sharlaimov   */
/* ArpSweep.cpp : Implementation of DLL Exports. */

#include "stdafx.h"
#include "resource.h"

// The module attribute causes DllMain, DllRegisterServer and DllUnregisterServer to be automatically implemented for you
[ module(dll, uuid = "{049F9861-0AC4-4FC3-9E7D-DEB15387682C}", 
		 name = "ArpSweep", 
		 helpstring = "ArpSweep 1.0",
		 resource_name = "IDR_ARPSWEEP") ]
class CArpSweepModule
{
public:
// Override CAtlDllModuleT members
};
		 
