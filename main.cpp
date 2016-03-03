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

template <typename Interface>
class IFooPtrT : public Com::Pointer<Interface>
{
public:
	IFooPtrT(Interface* value = nullptr)
		: Com::Pointer<Interface>(value)
	{
	}
	IFooPtrT<Interface>& operator=(Interface* value)
	{
		using Base = Com::Pointer<Interface>;
		Base::operator=(value);
		return *this;
	}
	operator Interface*() const
	{
		return p;
	}
	void Bar()
	{
		auto hr = p->Bar();
		Check(hr, __FUNCTION__, "");
	}
};
using IFooPtr = IFooPtrT<IFoo>;

template <>
class Com::TypeInfo<IFoo*>
{
public:
	using In = InValue<IFoo*, IFooPtr>;
	using InOut = InOutValue<IFoo*, IFooPtr>;
	using Retval = RetvalValue<IFooPtr, IFoo*>;
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

template <typename Argument> void Test(Argument argument) {}

void TestLong(long* pointer = nullptr)
{
	long value = 0;
	Test<long*>(Com::Get(value));
	Test<long>(Com::Put(value));
	Test<long*>(Com::PutRef(value));
	Test<long>(Com::In(*pointer));
	Test<long&>(Com::InOut(pointer));
	Com::Retval(pointer) = value;
}

void TestString(BSTR* pointer = nullptr)
{
	std::string value;
	Test<BSTR*>(Com::Get(value));
	Test<BSTR>(Com::Put(value));
	Test<BSTR*>(Com::PutRef(value));
	Test<std::string>(Com::In(*pointer));
	Test<std::string&>(Com::InOut(pointer));
	Com::Retval(pointer) = value;
}

void TestWideString()
{
	std::wstring value;
	Test<BSTR*>(Com::Get(value));
	Test<BSTR>(Com::Put(value));
	Test<BSTR*>(Com::PutRef(value));
	//No In/InOut/Retval (only bound to std::string)
}

void TestBool(VARIANT_BOOL* pointer = nullptr)
{
	bool value = false;
	Test<VARIANT_BOOL*>(Com::Get(value));
	Test<VARIANT_BOOL>(Com::Put(value));
	Test<VARIANT_BOOL*>(Com::PutRef(value));
	Test<bool>(Com::In(*pointer));
	Test<bool&>(Com::InOut(pointer));
	Com::Retval(pointer) = value;
}

void TestDate(DATE* pointer = nullptr)
{
	auto value = std::chrono::system_clock::now();
	Test<DATE*>(Com::Get(value));
	Test<DATE>(Com::Put(value));
	Test<DATE*>(Com::PutRef(value));
	Test<std::chrono::system_clock::time_point>(Com::In(*pointer));
	Test<std::chrono::system_clock::time_point&>(Com::InOut(pointer));
	Com::Retval(pointer) = value;
}

void TestVariant(VARIANT* pointer = nullptr)
{
	Com::Variant value;
	Test<VARIANT*>(Com::Get(value));
	Test<VARIANT>(Com::Put(value));
	Test<VARIANT*>(Com::PutRef(value));
	Test<Com::Variant>(Com::In(*pointer));
	Test<Com::Variant&>(Com::InOut(pointer));
	Com::Retval(pointer) = value;
}

void TestInterface(IUnknown** pointer = nullptr)
{
	Com::Pointer<IUnknown> value;
	Test<IUnknown**>(Com::Get(value));
	Test<IUnknown*>(Com::Put(value));
	Test<IUnknown**>(Com::PutRef(value));
	Test<Com::Pointer<IUnknown>>(Com::In(*pointer));
	Test<Com::Pointer<IUnknown>&>(Com::InOut(pointer));
	Com::Retval(pointer) = value;
}

void TestSmartPointer(IFoo** pointer = nullptr)
{
	IFooPtr value;
	Test<IFoo**>(Com::Get(value));
	Test<IFoo*>(Com::Put(value));
	Test<IFoo**>(Com::PutRef(value));
	Test<IFooPtr>(Com::In(*pointer));
	Test<IFooPtr&>(Com::InOut(pointer));
	Com::Retval(pointer) = value;
}

HRESULT TestError()
{
	try
	{
		Com::CheckError(E_FAIL, __FUNCTION__, "test");
	}
	catch (...)
	{
		return Com::HandleException();
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
	TestVariant();
	TestInterface();
	TestSmartPointer();
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
