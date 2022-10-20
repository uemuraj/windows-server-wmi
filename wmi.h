#pragma once

#include <Windows.h>
#include <comdef.h>
#include <wbemidl.h>
#include <iostream>
#include <system_error>

template<typename T>
using com_ptr_t = _com_ptr_t<_com_IIID<T, &__uuidof(T)>>;


namespace Wmi
{
	class Object
	{
		com_ptr_t<IWbemClassObject> m_object;

	public:
		Object(const com_ptr_t<IWbemClassObject> & object) : m_object(object)
		{}

		Object(const variant_t & value)
		{
			if (value.vt == VT_UNKNOWN)
			{
				m_object = value.punkVal;
			}
		}

		variant_t GetProperty(const wchar_t * name);
		void PutProperty(const wchar_t * name, variant_t && value);

		operator IWbemClassObject * () const
		{
			return m_object.GetInterfacePtr();
		}
	};

	class Service
	{
		com_ptr_t<IWbemServices> m_service;

	public:
		Service(const CLSID & rclsid, const wchar_t * resource);
		~Service() = default;

		Object GetClassObject(const wchar_t * objectPath);
		Object CreateInParams(const wchar_t * objectPath, const wchar_t * methodName);

		Object Exec(const wchar_t * objectPath, const wchar_t * methodName);
		Object Exec(const wchar_t * objectPath, const wchar_t * methodName, IWbemClassObject * inParams);
	};
}

std::wostream & operator<<(std::wostream & out, IWbemClassObject * object);
std::wostream & operator<<(std::wostream & out, const variant_t & value);
