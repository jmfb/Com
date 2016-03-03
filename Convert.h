#pragma once
#include <Windows.h>
#include <OleAuto.h>
#include <string>
#include <locale>
#include <codecvt>
#include <chrono>
#include <ctime>
#include "Pointer.h"

namespace Com
{
	template <typename Type>
	Type& CheckPointer(Type* value);

	template <typename Source, typename Target>
	inline void Assign(const Source& source, Target& target)
	{
		target = source;
	}

	template <typename Type, typename Source, typename Target>
	class GetValue
	{
	private:
		Target& target;
		Source source = TypeInfo<Target>::DefaultSource;

	public:
		GetValue(Target& value)
			: target(value)
		{
		}
		GetValue(const GetValue<Type, Source, Target>& rhs) = default;
		~GetValue()
		{
			Assign(source, target);
		}

		GetValue<Type, Source, Target>& operator=(const GetValue<Type, Source, Target>& rhs) = default;

		operator Type*()
		{
			return &source;
		}
	};

	template <typename Type, typename Source, typename Target>
	class PutValue
	{
	private:
		Target target;

	public:
		PutValue(const Source& value)
		{
			Assign(value, target);
		}
		PutValue(const PutValue<Type, Source, Target>& rhs) = default;
		~PutValue() = default;

		PutValue<Type, Source, Target>& operator=(const PutValue<Type, Source, Target>& rhs) = default;

		operator Type()
		{
			return target;
		}
	};

	template <typename Type, typename Source, typename Target>
	class PutRefValue
	{
	private:
		Source& source;
		Target target;

	public:
		PutRefValue(Source& value)
			: source(value)
		{
			Assign(source, target);
		}
		PutRefValue(const PutRefValue<Type, Source, Target>& rhs) = default;
		~PutRefValue()
		{
			Assign(target, source);
		}

		PutRefValue<Type, Source, Target>& operator=(const PutRefValue<Type, Source, Target>& rhs) = default;

		operator Type*()
		{
			return &target;
		}
	};

	template <typename Source, typename Target>
	class InValue
	{
	private:
		Target target;

	public:
		InValue(Source value)
		{
			Assign(value, target);
		}
		operator Target()
		{
			return target;
		}
	};

	template <typename Source, typename Target>
	class InOutValue
	{
	private:
		Source& source;
		Target target;

	public:
		InOutValue(Source* value)
			: source(CheckPointer(value))
		{
			Assign(source, target);
		}
		~InOutValue()
		{
			Assign(target, source);
		}
		operator Target&()
		{
			return target;
		}
	};

	template <typename Source, typename Target>
	class RetvalValue
	{
	private:
		Target& target;

	public:
		RetvalValue(Target* value)
			: target(CheckPointer(value))
		{
		}
		RetvalValue<Source, Target>& operator=(const Source& source)
		{
			Assign(source, target);
			return *this;
		}
	};

	template <typename Type>
	class TypeInfo
	{
	public:
		using Get = GetValue<Type, Type, Type>;
		using Put = Type;
		using PutRef = GetValue<Type, Type, Type>;
		using In = InValue<Type, Type>;
		using InOut = InOutValue<Type, Type>;
		using Retval = RetvalValue<Type, Type>;
	};

	template <typename Target>
	inline typename TypeInfo<Target>::Get Get(Target& value)
	{
		return{ value };
	}

	template <typename Source>
	inline typename TypeInfo<Source>::Put Put(const Source& source)
	{
		return{ source };
	}

	template <typename Source>
	inline typename TypeInfo<Source>::PutRef PutRef(Source& source)
	{
		return{ source };
	}

	template <typename Source>
	inline typename TypeInfo<Source>::In In(Source source)
	{
		return{ source };
	}

	template <typename Source>
	inline typename TypeInfo<Source>::InOut InOut(Source* source)
	{
		return{ source };
	}

	template <typename Target>
	inline typename TypeInfo<Target>::Retval Retval(Target* target)
	{
		return{ target };
	}

	////////////////////////////////////////////////////////////////////////////////
	// Intrinsic Type Support

	template <typename Target>
	class GetValue<Target, Target, Target>
	{
	private:
		Target& target;

	public:
		GetValue(Target& value)
			: target(value)
		{
		}
		GetValue(const GetValue<Target, Target, Target>& rhs) = default;
		~GetValue() = default;

		GetValue<Target, Target, Target>& operator=(const GetValue<Target, Target, Target>& rhs) = default;

		operator Target*()
		{
			return &target;
		}
	};

	template <typename Type>
	class InOutValue<Type, Type>
	{
	private:
		Type& source;

	public:
		InOutValue(Type* value)
			: source(CheckPointer(value))
		{
		}
		operator Type&()
		{
			return source;
		}
	};

	////////////////////////////////////////////////////////////////////////////////
	// Interfaces

	template <typename Interface>
	class GetInterface
	{
	private:
		Pointer<Interface>& target;
		Pointer<Interface> source;

	public:
		GetInterface(Pointer<Interface>& value)
			: target(value)
		{
		}
		~GetInterface()
		{
			target = source;
		}
		operator Interface**()
		{
			return &source;
		}
	};

	template <typename Interface>
	class PutInterface
	{
	private:
		Interface* source;

	public:
		PutInterface(Pointer<Interface> value)
			: source(value)
		{
		}
		operator Interface*()
		{
			return source;
		}
	};

	template <typename Interface>
	class PutRefInterface
	{
	private:
		Pointer<Interface>& source;

	public:
		PutRefInterface(Pointer<Interface>& value)
			: source(value)
		{
		}
		operator Interface**()
		{
			return &source.p;
		}
	};

	template <template <typename Interface> class SmartPointer, typename Interface>
	class TypeInfo<SmartPointer<Interface>>
	{
	public:
		using Get = GetInterface<Interface>;
		using Put = PutInterface<Interface>;
		using PutRef = PutRefInterface<Interface>;
	};

	template <typename Interface>
	class TypeInfo<Interface*>
	{
	public:
		using In = InValue<Interface*, Pointer<Interface>>;
		using InOut = InOutValue<Interface*, Pointer<Interface>>;
		using Retval = RetvalValue<Pointer<Interface>, Interface*>;
	};

	template <typename Interface>
	inline void Assign(const Interface*& source, Pointer<Interface>& target)
	{
		target = source;
	}

	template <typename Interface>
	inline void Assign(const Pointer<Interface>& source, Interface*& target)
	{
		if (target == source)
			return;
		if (target != nullptr)
			target->Release();
		target = source;
		if (target != nullptr)
			target->AddRef();
	}

	////////////////////////////////////////////////////////////////////////////////
	// std::string

	template <>
	class TypeInfo<std::string>
	{
	public:
		using Get = GetValue<BSTR, String, std::string>;
		using Put = PutValue<BSTR, std::string, String>;
		using PutRef = PutRefValue<BSTR, std::string, String>;
		static constexpr BSTR DefaultSource = nullptr;
	};

	template <>
	class TypeInfo<BSTR>
	{
	public:
		using In = InValue<BSTR, std::string>;
		using InOut = InOutValue<BSTR, std::string>;
		using Retval = RetvalValue<std::string, BSTR>;
	};

	template <>
	inline void Assign<BSTR, std::string>(const BSTR& source, std::string& target)
	{
		if (source == nullptr)
			target.clear();
		else
			target = std::wstring_convert<std::codecvt_utf8<wchar_t>>().to_bytes(source);
	}

	template <>
	inline void Assign<std::string, BSTR>(const std::string& source, BSTR& target)
	{
		if (target != nullptr)
			::SysFreeString(target);
		target = String{ source.c_str() }.Detach();
	}

	template <>
	inline void Assign<String, std::string>(const String& source, std::string& target)
	{
		Assign(static_cast<BSTR>(source), target);
	}

	template <>
	inline void Assign<std::string, String>(const std::string& source, String& target)
	{
		target = source.c_str();
	}

	////////////////////////////////////////////////////////////////////////////////
	// std::wstring

	template <>
	class TypeInfo<std::wstring>
	{
	public:
		using Get = GetValue<BSTR, String, std::wstring>;
		using Put = PutValue<BSTR, std::wstring, String>;
		using PutRef = PutRefValue<BSTR, std::wstring, String>;
		static constexpr BSTR DefaultSource = nullptr;
	};

	template <>
	inline void Assign<BSTR, std::wstring>(const BSTR& source, std::wstring& target)
	{
		if (source == nullptr)
			target.clear();
		else
			target = source;
	}

	template <>
	inline void Assign<std::wstring, BSTR>(const std::wstring& source, BSTR& target)
	{
		if (target != nullptr)
			::SysFreeString(target);
		target = ::SysAllocString(source.c_str());
	}

	template <>
	inline void Assign<String, std::wstring>(const String& source, std::wstring& target)
	{
		Assign(static_cast<BSTR>(source), target);
	}

	template <>
	inline void Assign<std::wstring, String>(const std::wstring& source, String& target)
	{
		target = source.c_str();
	}

	////////////////////////////////////////////////////////////////////////////////
	// bool

	template <>
	class TypeInfo<bool>
	{
	public:
		using Get = GetValue<VARIANT_BOOL, VARIANT_BOOL, bool>;
		using Put = PutValue<VARIANT_BOOL, bool, VARIANT_BOOL>;
		using PutRef = PutRefValue<VARIANT_BOOL, bool, VARIANT_BOOL>;
		static constexpr VARIANT_BOOL DefaultSource = VARIANT_FALSE;
	};

	template <>
	class TypeInfo<VARIANT_BOOL>
	{
	public:
		using In = InValue<VARIANT_BOOL, bool>;
		using InOut = InOutValue<VARIANT_BOOL, bool>;
		using Retval = RetvalValue<bool, VARIANT_BOOL>;
	};

	template <>
	inline void Assign<VARIANT_BOOL, bool>(const VARIANT_BOOL& source, bool& target)
	{
		target = source != VARIANT_FALSE;
	}

	template <>
	inline void Assign<bool, VARIANT_BOOL>(const bool& source, VARIANT_BOOL& target)
	{
		target = source ? VARIANT_TRUE : VARIANT_FALSE;
	}

	////////////////////////////////////////////////////////////////////////////////
	// std::chrono::time_point

	template <>
	class TypeInfo<std::chrono::system_clock::time_point>
	{
	public:
		using Get = GetValue<DATE, DATE, std::chrono::system_clock::time_point>;
		using Put = PutValue<DATE, std::chrono::system_clock::time_point, DATE>;
		using PutRef = PutRefValue<DATE, std::chrono::system_clock::time_point, DATE>;
		static constexpr DATE DefaultSource = 0;
	};

	template <>
	class TypeInfo<DATE>
	{
	public:
		using In = InValue<DATE, std::chrono::system_clock::time_point>;
		using InOut = InOutValue<DATE, std::chrono::system_clock::time_point>;
		using Retval = RetvalValue<std::chrono::system_clock::time_point, DATE>;
	};

	template <>
	inline void Assign<DATE, std::chrono::system_clock::time_point>(const DATE& source, std::chrono::system_clock::time_point& target)
	{
		SYSTEMTIME systemTime;
		::VariantTimeToSystemTime(source, &systemTime);
		std::tm tm;
		tm.tm_sec = systemTime.wSecond;
		tm.tm_min = systemTime.wMinute;
		tm.tm_hour = systemTime.wHour;
		tm.tm_mday = systemTime.wDay;
		tm.tm_mon = systemTime.wMonth - 1;
		tm.tm_year = systemTime.wYear - 1900;
		tm.tm_isdst = -1;
		tm.tm_wday = systemTime.wDayOfWeek;
		tm.tm_yday = -1;
		target = std::chrono::system_clock::from_time_t(std::mktime(&tm));
	}

	template <>
	inline void Assign<std::chrono::system_clock::time_point, DATE>(const std::chrono::system_clock::time_point& source, DATE& target)
	{
		auto time = std::chrono::system_clock::to_time_t(source);
		auto tm = *std::localtime(&time);
		SYSTEMTIME systemTime;
		systemTime.wSecond = tm.tm_sec;
		systemTime.wMinute = tm.tm_min;
		systemTime.wHour = tm.tm_hour;
		systemTime.wDay = tm.tm_mday;
		systemTime.wMonth = tm.tm_mon + 1;
		systemTime.wYear = tm.tm_year + 1900;
		systemTime.wDayOfWeek = tm.tm_wday;
		::SystemTimeToVariantTime(&systemTime, &target);
	}

	////////////////////////////////////////////////////////////////////////////////
	// Variant

	template <>
	class TypeInfo<Variant>
	{
	public:
		using Get = GetValue<VARIANT, Variant, Variant>;
		using Put = PutValue<VARIANT, Variant, Variant>;
		using PutRef = PutRefValue<VARIANT, Variant, Variant>;
		static constexpr VARTYPE DefaultSource = VT_EMPTY;
	};

	template <>
	class TypeInfo<VARIANT>
	{
	public:
		using In = InValue<VARIANT, Variant>;
		using InOut = InOutValue<VARIANT, Variant>;
		using Retval = RetvalValue<Variant, VARIANT>;
	};

	template <>
	inline void Assign<VARIANT, Variant>(const VARIANT& source, Variant& target)
	{
		target = source;
	}

	template <>
	inline void Assign<Variant, VARIANT>(const Variant& source, VARIANT& target)
	{
		auto hr = ::VariantCopy(&target, &source);
		if (FAILED(hr))
			throw std::runtime_error("VariantCopy failed.");
	}
}
