#include "wmi.h"

#define MACRO_CURRENT_LOCATION() __FILE__ "(" _CRT_STRINGIZE(__LINE__) ")"


namespace Wmi
{
	struct ComInitialize
	{
		ComInitialize()
		{
			if (auto hr = ::CoInitializeEx(nullptr, COINIT_MULTITHREADED); FAILED(hr))
			{
				throw std::system_error(hr, std::system_category(), MACRO_CURRENT_LOCATION());
			}
		}

		~ComInitialize()
		{
			::CoUninitialize();
		}
	};


	struct ComInitializeSecurity : ComInitialize
	{
		ComInitializeSecurity()
		{
			if (auto hr = ::CoInitializeSecurity(nullptr, -1, nullptr, nullptr, RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE, nullptr); FAILED(hr))
			{
				throw std::system_error(hr, std::system_category(), MACRO_CURRENT_LOCATION());
			}
		}

		~ComInitializeSecurity() = default;
	};


	struct wbem_category : public std::error_category
	{
		com_ptr_t<IWbemStatusCodeText> m_codeText;

		wbem_category() : m_codeText(__uuidof(WbemStatusCodeText))
		{}

		const char * name() const noexcept override
		{
			return "wbem error";
		}

		std::string message(int ev) const override
		{
			std::string msg;

			bstr_t text;

			if (auto hr = m_codeText->GetErrorCodeText(ev, 0, 0, text.GetAddress()); SUCCEEDED(hr))
			{
				if (int cch = ::WideCharToMultiByte(CP_THREAD_ACP, 0, text, text.length(), nullptr, 0, nullptr, nullptr); cch > 0)
				{
					msg.resize(cch);

					::WideCharToMultiByte(CP_THREAD_ACP, 0, text, text.length(), msg.data(), cch, nullptr, nullptr);
				}
			}

			return msg;
		}
	};


	Service::Service(const CLSID & rclsid, const wchar_t * resource)
	{
		thread_local static ComInitializeSecurity com;

		com_ptr_t<IWbemLocator> locator;

		if (auto hr = locator.CreateInstance(rclsid); FAILED(hr))
		{
			throw std::system_error(hr, wbem_category(), MACRO_CURRENT_LOCATION());
		}

		if (auto hr = locator->ConnectServer(bstr_t(resource), nullptr, nullptr, nullptr, 0, nullptr, nullptr, &m_service); FAILED(hr))
		{
			throw std::system_error(hr, wbem_category(), MACRO_CURRENT_LOCATION());
		}

		if (auto hr = ::CoSetProxyBlanket(m_service, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, nullptr, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE); FAILED(hr))
		{
			throw std::system_error(hr, std::system_category(), MACRO_CURRENT_LOCATION());
		}
	}

	Object Service::GetClassObject(const wchar_t * objectPath)
	{
		com_ptr_t<IWbemClassObject> object;

		if (auto hr = m_service->GetObject(bstr_t(objectPath), 0, nullptr, &object, nullptr); FAILED(hr))
		{
			throw std::system_error(hr, wbem_category(), MACRO_CURRENT_LOCATION());
		}

		return object;
	}

	Object Service::CreateInParams(const wchar_t * objectPath, const wchar_t * methodName)
	{
		com_ptr_t<IWbemClassObject> object;

		if (auto hr = m_service->GetObject(bstr_t(objectPath), 0, nullptr, &object, nullptr); FAILED(hr))
		{
			throw std::system_error(hr, wbem_category(), MACRO_CURRENT_LOCATION());
		}

		com_ptr_t<IWbemClassObject> inParams;

		if (auto hr = object->GetMethod(bstr_t(methodName), 0, &inParams, nullptr); FAILED(hr))
		{
			throw std::system_error(hr, wbem_category(), MACRO_CURRENT_LOCATION());
		}

		return inParams;
	}

	Object Service::Exec(const wchar_t * objectPath, const wchar_t * methodName)
	{
		com_ptr_t<IWbemClassObject> outParams;

		if (auto hr = m_service->ExecMethod(bstr_t(objectPath), bstr_t(methodName), 0, nullptr, CreateInParams(objectPath, methodName), &outParams, nullptr); FAILED(hr))
		{
			throw std::system_error(hr, wbem_category(), MACRO_CURRENT_LOCATION());
		}

		return outParams;
	}

	Object Service::Exec(const wchar_t * objectPath, const wchar_t * methodName, IWbemClassObject * inParams)
	{
		com_ptr_t<IWbemClassObject> outParams;

		if (auto hr = m_service->ExecMethod(bstr_t(objectPath), bstr_t(methodName), 0, nullptr, inParams, &outParams, nullptr); FAILED(hr))
		{
			throw std::system_error(hr, wbem_category(), MACRO_CURRENT_LOCATION());
		}

		return outParams;
	}

	variant_t Object::GetProperty(const wchar_t * name)
	{
		variant_t value;

		if (auto hr = m_object->Get(name, 0, &value, nullptr, nullptr); FAILED(hr))
		{
			throw std::system_error(hr, wbem_category(), MACRO_CURRENT_LOCATION());
		}

		return value;
	}

	void Object::PutProperty(const wchar_t * name, variant_t && value)
	{
		if (auto hr = m_object->Put(name, 0, &value, 0); FAILED(hr))
		{
			throw std::system_error(hr, wbem_category(), MACRO_CURRENT_LOCATION());
		}
	}
}


template<class T>
class SafeArrayData
{
	SAFEARRAY * m_psa;
	T * m_data;
	ULONG m_size;

public:
	T * begin()
	{
		return m_data;
	}

	T * end()
	{
		return m_data + m_size;
	}

	SafeArrayData(SAFEARRAY * psa) : m_psa(psa), m_data(nullptr), m_size(0)
	{
		VARTYPE vt{};

		if (auto hr = ::SafeArrayGetVartype(psa, &vt); FAILED(hr))
		{
			throw std::system_error(hr, std::system_category(), MACRO_CURRENT_LOCATION());
		}

		if (vt != GetVarType<T>())
		{
			throw std::system_error(E_UNEXPECTED, std::system_category(), MACRO_CURRENT_LOCATION());
		}

		LONG lbound{};

		if (auto hr = ::SafeArrayGetLBound(psa, 1, &lbound); FAILED(hr))
		{
			throw std::system_error(hr, std::system_category(), MACRO_CURRENT_LOCATION());
		}

		LONG ubound{};

		if (auto hr = ::SafeArrayGetUBound(psa, 1, &ubound); FAILED(hr))
		{
			throw std::system_error(hr, std::system_category(), MACRO_CURRENT_LOCATION());
		}

		if (lbound > ubound)
		{
			throw std::system_error(E_UNEXPECTED, std::system_category(), MACRO_CURRENT_LOCATION());
		}

		if (lbound < ubound)
		{
			if (auto hr = ::SafeArrayAccessData(psa, (void **) &m_data); FAILED(hr))
			{
				throw std::system_error(hr, std::system_category(), MACRO_CURRENT_LOCATION());
			}

			m_size = ubound - lbound;
		}
	}

	~SafeArrayData() noexcept
	{
		::SafeArrayUnaccessData(m_psa);
	}

protected:
	template<typename T>
	constexpr VARTYPE GetVarType() = delete;

	template<>
	constexpr VARTYPE GetVarType<BSTR>()
	{
		return VT_BSTR;
	}

	template<>
	constexpr VARTYPE GetVarType<IUnknown *>()
	{
		return VT_UNKNOWN;
	}
};

std::wostream & operator<<(std::wostream & out, IWbemClassObject * object)
{
	SAFEARRAY * psa{};

	if (auto hr = object->GetNames(nullptr, WBEM_FLAG_ALWAYS, nullptr, &psa); FAILED(hr))
	{
		throw std::system_error(hr, Wmi::wbem_category(), MACRO_CURRENT_LOCATION());
	}

	try
	{
		for (auto & name : SafeArrayData<BSTR>(psa))
		{
			variant_t value;

			if (auto hr = object->Get(name, 0, &value, nullptr, nullptr); FAILED(hr))
			{
				throw std::system_error(hr, Wmi::wbem_category(), MACRO_CURRENT_LOCATION());
			}

			out << name << L" = " << value << std::endl;
		}

		::SafeArrayDestroy(psa);
		return out;
	}
	catch (...)
	{
		::SafeArrayDestroy(psa);
		throw;
	}
}


std::wostream & operator<<(std::wostream & out, const variant_t & value)
{
	if (value.vt == VT_NULL)
	{
		return out << L"null";
	}

	if (value.vt == VT_BSTR)
	{
		return out << value.bstrVal;
	}

	if (value.vt == VT_UNKNOWN)
	{
		com_ptr_t<IWbemClassObject> object(value.punkVal);

		if (object)
		{
			return out << object.GetInterfacePtr();
		}
		else
		{
			return out << L"Unknown";
		}
	}

	if (value.vt & VT_ARRAY)
	{
		if (value.vt == (VT_ARRAY | VT_UNKNOWN))
		{
			int index = 0;

			for (auto & object : SafeArrayData<IUnknown *>(value.parray))
			{
				out << std::endl << L"---[" << index++ << L"]---" << std::endl << variant_t(object);
			}

			return out;
		}
	}

	variant_t dest;

	if (auto hr = ::VariantChangeTypeEx(&dest, &value, LOCALE_SYSTEM_DEFAULT, VARIANT_ALPHABOOL, VT_BSTR); SUCCEEDED(hr))
	{
		return out << dest.bstrVal;
	}

	return out << L"Type(" << value.vt << L')';
}
