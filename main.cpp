/*

    C++ WMI COM classes

    Visit https://github.com/g40

    Copyright (c) Jerry Evans, 2016-2024

    All rights reserved.

    The MIT License (MIT)

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in
    all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
    THE SOFTWARE.

*/

#include <Windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <conio.h>
#include <string.h>
#include <tchar.h>

#include <set>

//-----------------------------------------------------------------------------
#include <g40/nv2_opt.h>

//-----------------------------------------------------------------------------
#include "nv2_wmi.h"

//-----------------------------------------------------------------------------
// trace to stdout
#define TRMSG(args) std::wcout << args << std::endl;

//-----------------------------------------------------------------------------
#ifdef _UNICODE
int _tmain(int argc, char_t* argv[])
#else
int main(int argc, char* argv[])
#endif
{
    int ret = -1;
    try
    {
        //---------------------------------------------------------------------
        //
        nv2::throw_if(!uw32::IsProcessElevated(),
                    nv2::acc("This application requires administrative privileges. Please run as Administrator."));

        //---------------------------------------------------------------------
        bool help = false;
        bool test_enumeration = false;
        bool test_properties = false;
        bool test_methods = false;
        bool list_properties = false;
        bool list_methods = false;
        std::wstring targetName;
        // map options to default values
        std::vector<nv2::ap::Opt> opts = 
        {
            { _T("-?"), help, _T("Display help text") },
            { _T("--help"), help, _T("Display help text") },
            { _T("-te"), test_enumeration, _T("Test enumeration: Display Verbose WMI listing") },
            { _T("-tp"), test_properties, _T("Test properties (i.e. retrieve values)") },
            { _T("-tm"), test_methods, _T("Test methods") },
            { _T("-tn"), targetName, _T("Limit enumeration to classes that match. Use of '%' wildcard works..") },
            { _W("-lp"), list_properties, _W("List properties while enumerating.")},
            { _W("-lm"), list_methods, _W("List methods (and parameters) while enumerating.")},
        };

        // parse the command line. returns any positionals in vp
        std::vector<string_t> vp = nv2::ap::parse(argc, (const char_t**)argv, opts);

        if (help) {
            string_t s = nv2::ap::to_string(opts, _T("wmipp. Simple C++/WMI interop driver."));
            std::wcout << nv2::t2w(s);
            return 0;
        }

        //
        if (!targetName.empty()) {
            test_enumeration = true;
        }
        //
        nv2::wmi::ComInit ci;
        // 
        nv2::wmi::Services srv;
        //
        if (test_enumeration)
        {
            // enumerate all classes in the namespace           
            //std::set<std::wstring> classNames = srv.GetClassNames();
            std::set<std::wstring> classNames = srv.GetClassNames(targetName.c_str());
            //std::set<std::wstring> classNames = srv.GetClassNames(_W("Win32_Processor"));
            // dump all properties and methods
            for (auto& className : classNames) 
            {
                TRMSG("Classname: " << className);
                nv2::wmi::Object obj = srv.GetObject(className);

                if (list_properties) {
                    std::vector<std::wstring> props = obj.GetProperties();
                    for (auto& prop : props)
                    {
                        TRMSG("\tProperty: " << prop);
                    }
                }

                if (list_methods) 
                {
                    std::vector<nv2::wmi::MethodDef> methods = obj.GetMethods();
                    for (auto& method : methods)
                    {
                        TRMSG("\tMethod: " << method.name);
                        for (auto& ipp : method.ipParams)
                            TRMSG("\t\tIn: " << ipp);
                        for (auto& opp : method.opParams)
                            TRMSG("\t\tOut: " << opp);
                    }
                }
            }
        }

        if (test_properties)
        {
            std::wstring key = _W("Win32_LogicalDisk");
            // nope. not on Win10
            //key = _W("MSAcpi_ThermalZoneTemperature");
            //
            //key = _W("Win32_Processor");
            //
            std::vector<nv2::wmi::Object> disks = srv.GetInstances(key);
            for (auto& disk : disks)
            {
                std::vector<std::wstring> props = disk.GetProperties();
                for (auto& prop : props)
                {
                    std::wstring sv = disk.GetValue(prop);
                    std::wcout << prop << " => " << sv << std::endl;
                }
            }
        }
      
        // use WMI to actually do something.
        if (test_methods)
        {
            std::wstring key = _W("Win32_LogicalDisk");
            std::vector<nv2::wmi::Object> disks = srv.GetInstances(key);
            for (auto& disk : disks)
            {
                std::wstring diskId = disk.GetValue(_W("DeviceID"));
                if (diskId == _W("G:"))
                {
                    // see next snippet
                    int ir = -1;
                    nv2::wmi::Object::param_list iparams = {
                        { _T("FixErrors"), false },
                        { _T("OKToRunAtBootUp"), false },
                    };
                    //
                    nv2::wmi::Object::param_map oparams;
                    //
                    CComVariant result = disk.ExecMethod(_W("chkdsk"),
                                                        iparams, 
                                                        oparams);
                    //
                    ir = result.intVal;
                }
            }
        }
        //
        ret = 0;
    }
    catch (const std::exception& ex)
    {
        std::cout << "Error: " << ex.what() << std::endl;
    }
    catch (const DWORD& ex)
    {
        std::wcout << "Error: " << nv2::s_error(ex).c_str() << std::endl;
    }
    catch (...)
    {
        std::cout << "Unknown error ..." << std::endl;
    }
    //
    return ret;
}
