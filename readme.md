### wmipp ###

C++ classes for Windows WMI COM API

#### A quick example ####

```
//-----------------------------------------------------------------------------
// ensure the right header is included.
#include "nv2_wmi.h"
//
try 
{
	// ensure COM is initialized
	nv2::wmi::ComInit ci;
	// handle the service setup
	nv2::wmi::WMIServices srv;
	// Disks ...
	std::wstring key = _W("Win32_LogicalDisk");
	// find all disks
	std::vector<nv2::wmi::Object> disks = srv.GetInstances(key);
	for (auto& disk : disks)
	{
		// dump all of the properties of the disk instance
		std::vector<std::wstring> props = disk.GetProperties();
		for (auto& prop : props)
		{
			std::wstring sv = obj.GetValue(prop);
			std::wcout << prop << " => " << sv << std::endl;
		}
	}
}
catch(...)
{
	// exception!
}

```

#### Calling WMI instance methods ####


```
	std::wstring key = _W("Win32_LogicalDisk");
	std::vector<nv2::wmi::Object> disks = srv.GetInstances(key);
	for (auto& disk : disks)
	{
		std::wstring diskId = disk.GetValue(_W("DeviceID"));
		if (diskId == _W("G:"))
		{
			// as WMI methods can have varying arity
			// makes sense to use a type-safe list
			nv2::wmi::Object::param_list iparams = {
				{ _T("FixErrors"), false },
				{ _T("OKToRunAtBootUp"), false },
			};
			// optional
			nv2::wmi::Object::param_map oparams;
			//
			CComVariant result = disk.ExecMethod(_W("chkdsk"),
												iparams, 
												oparams);
			//
			std::cout << "chkdsk returned " << result.intVal << std::endl;
		}
	}
```
