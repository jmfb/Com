#define _CRT_SECURE_NO_WARNINGS
#include "Com.h"

struct __declspec(uuid("4DE37185-5855-4B6F-A0D8-F4B530548510")) IFoo : public IDispatch
{
	virtual HRESULT __stdcall Bar() = 0;
};

struct __declspec(uuid("4DE37185-5855-4B6F-A0D8-F4B530548511")) IFoo2 : public IDispatch
{
	virtual HRESULT __stdcall Bar2() = 0;
};

extern const GUID CLSID_Foo = { 0x9627fa9f, 0x699b, 0x4afd, { 0x84, 0x86, 0xcc, 0xd8, 0x4, 0x47, 0x9b, 0x5 } };

class Foo : public Com::Object<Foo, &CLSID_Foo, IFoo, IFoo2>
{
public:
	HRESULT __stdcall Bar() final
	{
		return S_OK;
	}
	HRESULT __stdcall Bar2() final
	{
		return S_OK;
	}
};

void LongOutput(long* result)
{
}
void LongInput(long value)
{
}
void StringOutput(BSTR* result)
{
}
void StringInput(BSTR value)
{
}
void BoolOutput(VARIANT_BOOL* result)
{
}
void BoolInput(VARIANT_BOOL value)
{
}
void DateOutput(DATE* result)
{
}
void DateInput(DATE value)
{
}

void TestLong()
{
	long value = 0;
	LongOutput(Com::Get(value));
	LongInput(Com::Put(value));
	LongOutput(Com::PutRef(value));
}

void TestString()
{
	std::string value;
	StringOutput(Com::Get(value));
	StringInput(Com::Put(value));
	StringOutput(Com::PutRef(value));
}

void TestWideString()
{
	std::wstring value;
	StringOutput(Com::Get(value));
	StringInput(Com::Put(value));
	StringOutput(Com::PutRef(value));
}

void TestBool()
{
	bool value = false;
	BoolOutput(Com::Get(value));
	BoolInput(Com::Put(value));
	BoolOutput(Com::PutRef(value));
}

void TestDate()
{
	auto value = std::chrono::system_clock::now();
	DateOutput(Com::Get(value));
	DateInput(Com::Put(value));
	DateOutput(Com::PutRef(value));
}

HRESULT TestError()
{
	try
	{
		Com::CheckError(E_FAIL, __FUNCTION__, "test");
	}
	catch (const Com::Error& error)
	{
		return error.SetErrorInfo();
	}
	return S_OK;
}

int main()
{
	TestLong();
	TestString();
	TestWideString();
	TestBool();
	TestDate();
	TestError();
	return 0;
}

extern "C" BOOL __stdcall DllMain(HINSTANCE instance, DWORD reason, void* reserved)
{
	if (reason == DLL_PROCESS_ATTACH)
		Com::Module::GetInstance().Initialize(instance);
	return TRUE;
}

HRESULT __stdcall DllCanUnloadNow()
{
	return Com::Module::GetInstance().CanUnload() ? S_OK : S_FALSE;
}

HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppvObject)
{
	return Com::ObjectList<
		Foo
	>::Create(rclsid, riid, ppvObject);
}
