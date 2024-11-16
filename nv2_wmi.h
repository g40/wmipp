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


#pragma once

#include <map>
#include <vector>
#include <initializer_list>
#include <string>
#include <cstdint>
#include <atlbase.h>
#include <atlcom.h>
#include <comdef.h>
#include <Wbemidl.h>

#pragma comment(lib, "wbemuuid.lib")

#include <g40/nv2_w32.h>

namespace nv2
{
	namespace wmi
	{
		//-------------------------------------------------------------------------
		// COM initializer
		// https://learn.microsoft.com/en-us/windows/win32/com/com-error-codes-1
		class ComInit
		{
			bool ok = false;
		public:
			ComInit()
			{
				// initialize COM (must do before InitializeForBackup works)
				HRESULT result = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
				uw32::trace_hresult(LFL "ComInit", result);
				nv2::throw_if(result != S_OK, nv2::acc(LFL));
				ok = true;
			}
			~ComInit()
			{
				if (ok) 
				{
					CoUninitialize();
					ok = false;
				}
			}
		};

		//-----------------------------------------------------------------------------
		// for SAFEARRAY cleanup
		class SAHandle 
		{
			SAFEARRAY* m_psa = nullptr;
		
		public:
			
			explicit SAHandle(SAFEARRAY* psa) : m_psa(psa) {}

			~SAHandle() 
			{
				if (m_psa) {
					HRESULT hr = SafeArrayDestroy(m_psa);
					nv2::throw_if(hr != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hr))));
					m_psa = nullptr;
				}
			}
		};

		//-----------------------------------------------------------------------------
		// package up all method attributes
		struct MethodDef
		{
			// method name
			std::wstring name;
			// input parameter names
			std::vector<std::wstring> ipParams;
			// output (reference) parameter names
			std::vector<std::wstring> opParams;
		};

		//-----------------------------------------------------------------------------
		// get the collection of names associated with the IWbemClassObject
		static
		std::vector<std::wstring>
			EnumNames(const CComPtr<IWbemClassObject>& pClass)
		{
			std::vector<std::wstring> ret;

			SAFEARRAY* psaNames = NULL;
			HRESULT hResult = pClass->GetNames(
				NULL,
				WBEM_FLAG_ALWAYS | WBEM_FLAG_NONSYSTEM_ONLY,
				NULL,
				&psaNames);
			nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));
			
			// handle cleanup
			SAHandle sah(psaNames);	

			// Get the number of properties.
			long lLower = 0;
			hResult = SafeArrayGetLBound(psaNames, 1, &lLower);
			nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));

			long lUpper = 0;
			hResult = SafeArrayGetUBound(psaNames, 1, &lUpper);
			nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));

			for (long i = lLower; i <= lUpper; i++)
			{
				CComBSTR name;
				hResult = SafeArrayGetElement(
					psaNames,
					&i,
					&name);
				nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));
				ret.push_back((BSTR)name);
			}
			//
			return ret;
		}

		//-----------------------------------------------------------------------------
		// badly named 
		class Object
		{
		// protected:
				
			CComPtr<IWbemClassObject> m_pObj;
			CComPtr<IWbemServices> m_pServices;

		public:

			// Creates an empty (invalid) WMIObject
			// Object() {}

			//-----------------------------------------------------------------------------
			// Creates a WMIObject for a given IWbemClassObject instance
			Object(const CComPtr<IWbemClassObject>& pObj, 
					const CComPtr<IWbemServices>& pServices)
				: m_pObj(pObj)
				, m_pServices(pServices)
			{
			}

			//-----------------------------------------------------------------------------
			Object(const VARIANT* pInterface, const CComPtr<IWbemServices>& pServices)
			{
				nv2::throw_if(pInterface == nullptr, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(E_POINTER))));
				nv2::throw_if(pInterface && pInterface->vt != VT_UNKNOWN, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(E_INVALIDARG))));

				nv2::throw_if(!pServices, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(E_POINTER))));

				pInterface->punkVal->QueryInterface(&m_pObj);

				m_pServices = pServices;
			}

			//-----------------------------------------------------------------------------
			Object(IUnknown* pUnknown, const CComPtr<IWbemServices>& pServices)
			{
				nv2::throw_if(!pUnknown, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(E_POINTER))));
				pUnknown->QueryInterface(&m_pObj);

				m_pServices = pServices;
			}

			//-----------------------------------------------------------------------------
			// Tests WMIObject instance for validity
			bool Valid() const
			{
				return m_pObj != NULL;
			}

			//-----------------------------------------------------------------------------
			// get all properties
			std::vector<std::wstring> 
				GetProperties() const
			{
				std::vector<std::wstring> ret;
				//
				SAFEARRAY* psaNames = NULL;
				//
				HRESULT hResult = m_pObj->GetNames(NULL,
													WBEM_FLAG_ALWAYS | WBEM_FLAG_NONSYSTEM_ONLY,
													NULL,
													&psaNames);
				nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));
				// for cleanup
				SAHandle sah(psaNames);
				//
				ret = EnumNames(m_pObj);
				//
				return ret;
			}

			//-----------------------------------------------------------------------------
			std::wstring 
				GetValue(const std::wstring& property) const
			{
				nv2::throw_if(!m_pObj, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(E_POINTER))));

				CComVariant val;
				HRESULT hResult = m_pObj->Get(property.c_str(), 0, &val, NULL, NULL);
				nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));

				std::wstring ret;
				if (val.vt == VT_NULL) 
				{
					ret = _W("NULL");
				}
				else if (val.vt == VT_BOOL)
				{
					ret = (val.bVal ? _W("true") : _W("false"));
				}
				else if (!((val.vt == VT_LPWSTR) || (val.vt == VT_BSTR)))
				{	
					val.ChangeType(VT_BSTR);
					ret = val.bstrVal;
				}
				else
				{
					ret = val.bstrVal;
				}
				return ret;
			}

			//-----------------------------------------------------------------------------
			int
				GetIValue(const std::wstring& property) const
			{
				nv2::throw_if(!m_pObj, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(E_POINTER))));

				CComVariant val;
				HRESULT hResult = m_pObj->Get(property.c_str(), 0, &val, NULL, NULL);
				nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));

				// should be a trait
				if (val.vt != VT_INT)
					val.ChangeType(VT_INT);

				//
				return val.intVal;
			}

			//-----------------------------------------------------------------------------
			std::vector<MethodDef> GetMethods()
			{
				long lFlags = 0;
				std::vector<MethodDef> ret;

				CComPtr<IWbemClassObject> pClassDefinition;
				//std::wstring className = asString(L"__CLASS");
				std::wstring className = GetValue(L"__CLASS");
				HRESULT hR = m_pServices->GetObject(CComBSTR(className.c_str()), 0, NULL, &pClassDefinition, NULL);
				nv2::throw_if(hR != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hR))));

				HRESULT hr = pClassDefinition->BeginMethodEnumeration(lFlags);
				if (SUCCEEDED(hr))
				{
					for (;;)
					{
						// 
						CComBSTR name;
						CComPtr<IWbemClassObject> ppInSignature;
						CComPtr<IWbemClassObject> ppOutSignature;
						hr = pClassDefinition->NextMethod(0, &name, &ppInSignature, &ppOutSignature);
						if (FAILED(hr)) {
							break;
						}
						//
						if (name == nullptr)
							break;
						//
						MethodDef def;
						def.name = (BSTR)name;
						if (ppInSignature)
							def.ipParams = EnumNames(ppInSignature);
						if (ppOutSignature)
							def.opParams = EnumNames(ppOutSignature);
						//
						ret.push_back(def);
					}
					hr = pClassDefinition->EndMethodEnumeration();
				}
				else
				{
					DBMSG(nv2::s_error(uw32::Win32FromHResult(hr)));
				}
				return ret;	
			}

			//-----------------------------------------------------------------------------
			// simplify parameter declarations
			using name_t = std::wstring;
			using value_t = CComVariant;
			using param_t = std::pair<name_t, value_t>;
			using param_list = std::initializer_list<param_t>;
			using param_map = std::map<name_t,value_t>;
			using param_iterator = param_map::iterator;
			//-----------------------------------------------------------------------------
			// nonstandard extension
#pragma warning ( suppress: 4239 )
			CComVariant
				ExecMethod(LPCWSTR lpMethodName,
					const param_list& iparams,
					param_map& oparams = param_map())
			{
				nv2::throw_if(m_pServices == NULL, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(E_POINTER))));
				nv2::throw_if(m_pObj == NULL, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(E_POINTER))));
				
				CComPtr<IWbemClassObject> pClassDefinition;
				std::wstring className = GetValue(L"__CLASS");
				HRESULT hR = m_pServices->GetObject(CComBSTR(className.c_str()), 0, NULL, &pClassDefinition, NULL);
				nv2::throw_if(hR != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hR))));

				CComPtr<IWbemClassObject> pInParamsClass;
				hR = pClassDefinition->GetMethod(lpMethodName, 0, &pInParamsClass, NULL);

				//if (pInParamsClass) {
				CComPtr<IWbemClassObject> pInParamsInstance;
				hR = pInParamsClass->SpawnInstance(0, &pInParamsInstance);
				// check hR
				DBMSG("pInParamsClass->SpawnInstance => " << nv2::s_error(uw32::Win32FromHResult(hR)));
				//
				for (auto& iparam : iparams)
				{
					hR = pInParamsInstance->Put(iparam.first.c_str(), 
												0, 
													&const_cast<CComVariant&>(iparam.second), 
												iparam.second.vt);
					// check hR
					DBMSG(iparam.first << " => " << nv2::s_error(uw32::Win32FromHResult(hR))); 
				}

				// synchronous call	
				long lFlags = 0;
				IWbemContext* pCtx = NULL;
				CComPtr<IWbemClassObject> pOutParamsInstance;

				// https://learn.microsoft.com/en-us/windows/win32/wmisdk/describing-a-class-object-path
				std::wstring relPath = GetValue(L"__RELPATH");
				hR = m_pServices->ExecMethod(CComBSTR(relPath.c_str()), 
					CComBSTR(lpMethodName), 
					lFlags,		// synchronous
					pCtx,	// 
					pInParamsInstance, 
					&pOutParamsInstance, 
					NULL);
				DBMSG("m_pServices->ExecMethod => " << nv2::s_error(uw32::Win32FromHResult(hR)));

				CComVariant ret;
				// always get the return value
				if (pOutParamsInstance) 
				{
					// there is always a return value
					hR = pOutParamsInstance->Get(_W("ReturnValue"), 0, &ret, NULL, NULL);
					DBMSG("pOutParamsInstance->Get => " << nv2::s_error(uw32::Win32FromHResult(hR)));

					hR = pOutParamsInstance->BeginEnumeration(WBEM_FLAG_NONSYSTEM_ONLY);
					DBMSG("pOutParamsInstance->BeginEnumeration => " << nv2::s_error(uw32::Win32FromHResult(hR)));
					for (;;)
					{
						CComBSTR name;
						CComVariant val;
						// name.AssignBSTR(NULL);
						hR = pOutParamsInstance->Next(0, &name, &val, NULL, NULL);
						if (!name)
							break;
						if (name == _W("ReturnValue"))	
							continue;
						// otherwise map it
						oparams[std::wstring(name)] = val;
					}
					hR = pOutParamsInstance->EndEnumeration();
					DBMSG("pOutParamsInstance->EndEnumeration => " << nv2::s_error(uw32::Win32FromHResult(hR)));
				}
				//
				return ret;
			}
		};	// WMIObject

		//---------------------------------------------------------------------
		// Represents a root WMI Services object
		class Services
		{
		private:

			CComPtr<IWbemServices> m_pService;

		public:

			//-----------------------------------------------------------------
			Services(LPCWSTR lpResourcePath = L"ROOT\\CIMV2") throw()
			{
				HRESULT hResult = CoInitializeSecurity(NULL,
					-1,                          // COM authentication
					NULL,                        // Authentication services
					NULL,                        // Reserved
					RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
					RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
					NULL,                        // Authentication info
					EOAC_NONE,                   // Additional capabilities 
					NULL                         // Reserved
				);

				nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));

				CComPtr<IWbemLocator> pLocator;
				hResult = pLocator.CoCreateInstance(CLSID_WbemLocator);
				nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));

				hResult = pLocator->ConnectServer(CComBSTR(lpResourcePath), NULL, NULL, NULL, 0, NULL, NULL, &m_pService);
				nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));

				hResult = CoSetProxyBlanket(m_pService, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, NULL, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE);
				nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));
			
			}

			//---------------------------------------------------------------------
			// Tests WMIServices for validity
			bool Valid() const
			{
				return m_pService != NULL;
			}

			//-----------------------------------------------------------------------------
			// Returns a set of all instances of a given object.
			std::vector<Object>
				GetInstances(LPCWSTR lpClassName) const
			{
				int Flags = 0;
				IWbemContext* pCtx = NULL;

				nv2::throw_if(m_pService == NULL, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(E_POINTER))));

				CComPtr<IEnumWbemClassObject> pEnum;
				HRESULT hResult = m_pService->CreateInstanceEnum(CComBSTR(lpClassName), Flags, pCtx, &pEnum);
				nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));

				std::vector<Object> ret;
				// ObjectCollection(pEnum, m_pService);
				while (pEnum)
				{
					ULONG returned = 0;
					CComPtr<IWbemClassObject> pObject;
					hResult = pEnum->Next(0, 1, &pObject, &returned);
					if (returned == 0)
						break;
					nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));
					//
					Object obj(pObject, m_pService);
					//
					ret.push_back(obj);
				}
				return ret;
			}

			//-----------------------------------------------------------------------------
			std::vector<Object>
				GetInstances(const std::wstring& className) const
			{
				return GetInstances(className.c_str());
			}

			//-----------------------------------------------------------------------------
			// Returns a WMI Object representing a given WMI object, or a class
			Object 
			GetObject(LPCWSTR lpObjectName)
			{
				int Flags = 0;
				IWbemContext* pCtx = NULL;
				//
				nv2::throw_if(m_pService == NULL, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(E_POINTER))));

				CComPtr<IWbemClassObject> pObj;

				HRESULT hResult = m_pService->GetObject(CComBSTR(lpObjectName), Flags, pCtx, &pObj, NULL);
				nv2::throw_if(hResult != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hResult))));
					
				return Object(pObj, m_pService);
			}		

			//-----------------------------------------------------------------------------
			//
			Object
				GetObject(const std::wstring& objectName)
			{
				Object ret = GetObject(objectName.c_str());
				return ret;
			}

			//-----------------------------------------------------------------------------
			// enumerate all class names in this namespace
			std::set<std::wstring> 
				GetClassNames(const std::wstring& filter = _W(""))
			{
				//
				std::wstring query = _W("SELECT * FROM meta_class");
				if (!filter.empty()) {
					query += _W(" where __CLASS LIKE ");
					query += _W("'");
					query += filter;
					query += _W("'");
				}
				std::set<std::wstring> ret;
				CComPtr<IEnumWbemClassObject> enumerator;
				HRESULT hres = m_pService->ExecQuery(bstr_t("WQL"),
					bstr_t(query.c_str()),
					WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
					NULL,
					&enumerator);
				nv2::throw_if(hres != S_OK, (nv2::acc(LFL) << nv2::s_error(uw32::Win32FromHResult(hres))));

				while (enumerator) 
				{
					ULONG uReturn = 0;
					CComPtr<IWbemClassObject> pClass;
					//
					hres = enumerator->Next(WBEM_INFINITE, 1, &pClass, &uReturn);
					if (uReturn == 0) 
					{
						break;
					}

					// https://learn.microsoft.com/en-us/windows/win32/api/wbemcli/nf-wbemcli-iwbemclassobject-get?devlangs=cpp&f1url=%3FappId%3DDev17IDEF1%26l%3DEN-US%26k%3Dk(ATLCOMCLI%252FATL%253A%253A_NoAddRefReleaseOnCComPtr%253A%253AGet)%3Bk(ATL%253A%253A_NoAddRefReleaseOnCComPtr%253A%253AGet)%3Bk(WBEMCLI%252FIWbemClassObject%253A%253AGet)%3Bk(UNKNWNBASE%252FIUnknown%253A%253AGet)%3Bk(IWbemClassObject%253A%253AGet)%3Bk(IUnknown%253A%253AGet)%3Bk(Get)%3Bk(DevLang-C%252B%252B)%3Bk(TargetOS-Windows)%26rd%3Dtrue
					// get the class name from pClass
					CComVariant vtClassName;
					hres = pClass->Get(L"__CLASS", 0, &vtClassName, 0, 0);
					if (SUCCEEDED(hres)) 
					{
						ret.insert(vtClassName.bstrVal);
						}
					// almost never present/available so we ignore
					// CComVariant vtDesc;
				}
				//
				return ret;
			}
		};

		//-------------------------------------------
		static
			void
			test()
		{
			if (0)
			{
				ComInit ci;
				// WMIServices srv(L"root\\WMI");
				Services srv;
				//
				std::wstring key = _W("Win32_LogicalDisk");
				std::vector<Object> objs = srv.GetInstances(key);
				for (auto& obj : objs)
				{
					//
					std::wstring did = obj.GetValue(L"DeviceID");
					std::wstring vid = obj.GetValue(L"VolumeName");
					std::wstring vsn = obj.GetValue(L"VolumeSerialNumber");
					int blockSize = obj.GetIValue(L"BlockSize");
					//int blockCount = obj.GetIValue(L"NumberOfBlocks");
					//int compressed = obj.GetIValue(L"Compressed");
					//
					std::wcout << "test: " << did << " " << vid << " " << vsn << " " << blockSize;

					if (did == L"G:")
					{
						// 
						int ir = -1;
						//
						nv2::wmi::Object::param_list iparams = 
						{
							{ _T("FixErrors"), false },
							{ _T("OKToRunAtBootUp"), false },
						};
						//
						nv2::wmi::Object::param_map oparams;
						//
						CComVariant result = obj.ExecMethod(L"chkdsk", iparams, oparams);
						//
						ir = result.intVal;
					}
				}
			}
		}
	}
}

