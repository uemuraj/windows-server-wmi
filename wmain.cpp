#include <Windows.h>
#include <locale>
#include <iostream>
#include <system_error>

#include "wmi.h"

int wmain(int argc, wchar_t * argv)
{
	try
	{
		std::locale::global(std::locale(""));

#if 0
		//
		// PS> (Invoke-WmiMethod -Namespace "ROOT/Microsoft/Windows/ServerManager" -Path "MSFT_ServerManagerTasks" -Name "GetServerFeature").cmdletOutput
		//
		Wmi::Service manager(__uuidof(WbemLocator), LR"(ROOT/Microsoft/Windows/ServerManager)");

		auto result = manager.Exec(L"MSFT_ServerManagerTasks", L"GetServerFeature");

		std::wcout << result.GetProperty(L"cmdletOutput");

		return 0;
#else
		Wmi::Service manager(__uuidof(WbemLocator), LR"(ROOT/Microsoft/Windows/ServerManager)");

		auto inParams = manager.CreateInParams(L"MSFT_ServerManagerDeploymentTasks", L"GetServerComponentsAsync");
		auto inParam1 = manager.GetClassObject(L"MSFT_ServerManagerRequestGuid");

		inParams.PutProperty(L"RequestGuid", variant_t(inParam1));

		auto result = manager.Exec(L"MSFT_ServerManagerDeploymentTasks", L"GetServerComponentsAsync", inParams);

		Wmi::Object object = result.GetProperty(L"EnumerationState");

		for (;;)
		{
			if (BYTE state = object.GetProperty(L"RequestState"); state == 0)
			{
				::Sleep(1000);
				continue;
			}
			else if (state != 1)
			{
				std::wcout << result.GetProperty(L"Error");
				break;
			}
			else
			{
				std::wcout << result.GetProperty(L"ServerComponents");
				break;
			}
		}

		return result.GetProperty(L"ReturnValue");
#endif
	}
	catch (const std::system_error & e)
	{
		std::cerr << e.what() << std::endl;

		return e.code().value();
	}
	catch (const std::exception & e)
	{
		std::cerr << e.what() << std::endl;

		return -1;
	}
}
