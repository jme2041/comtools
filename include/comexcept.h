// comexcept.h ////////////////////////////////////////////////////////////////
//
// ComTools::ComException: COM GetErrorInfo() Wrapper
//
// ComTools::ComException is released under the MIT license.
//
// Copyright 2022, Jeffrey M. Engelmann
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to
// deal in the Software without restriction, including without limitation the
// rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
// sell copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
// IN THE SOFTWARE.
//

#ifndef COMEXCEPT_H
#define COMEXCEPT_H

#include <Windows.h>
#include <utility>
#include <string>
#include <type_traits>
#include "ubstr.h"

namespace ComTools {

    class ComException {
        // Note that ComException does not derive from std::exception because
        // the ComException copy constructor is not noexcept

        HRESULT m_hr = 0;
        std::wstring m_source;
        std::wstring m_description;
        std::wstring m_help_file;
        DWORD m_help_context = 0;
        GUID m_guid = IID_NULL;

    public:
        friend void swap(ComException& a, ComException& b) noexcept
        {
            std::swap(a.m_hr, b.m_hr);
            std::swap(a.m_source, b.m_source);
            std::swap(a.m_description, b.m_description);
            std::swap(a.m_help_file, b.m_help_file);
            std::swap(a.m_help_context, b.m_help_context);
            std::swap(a.m_guid, b.m_guid);
        }

        ComException()
            noexcept(std::is_nothrow_default_constructible<std::wstring>::value) = default;

        ComException(HRESULT const hr) :
            m_hr(hr),
            m_source(),
            m_description(),
            m_help_file(),
            m_help_context(0),
            m_guid()
        {
            IErrorInfo* pei = nullptr;
            HRESULT hr2 = GetErrorInfo(0, &pei);
            if (hr2 == S_OK && pei)
            {
                UBSTR source;
                if (SUCCEEDED(pei->GetSource(set(source)))) m_source = source.to_wstring();

                UBSTR description;
                if (SUCCEEDED(pei->GetDescription(set(description)))) m_description = description.to_wstring();

                UBSTR help_file;
                if (SUCCEEDED(pei->GetHelpFile(set(help_file)))) m_help_file = help_file.to_wstring();

                DWORD help_context{};
                if (SUCCEEDED(pei->GetHelpContext(&help_context))) m_help_context = help_context;

                GUID guid{};
                if (SUCCEEDED(pei->GetGUID(&guid))) m_guid = guid;

                pei->Release();
            }
        }

        virtual ~ComException() noexcept { };

        ComException(ComException const& obj) :
            m_hr(obj.m_hr),
            m_source(obj.m_source),
            m_description(obj.m_description),
            m_help_file(obj.m_help_file),
            m_help_context(obj.m_help_context),
            m_guid(obj.m_guid) { }

        ComException(ComException&& obj)
            noexcept(std::is_nothrow_default_constructible<std::wstring>::value) :
            ComException()
        {
            swap(*this, obj);
        }

        ComException& operator=(ComException obj)
        {
            swap(*this, obj);
            return *this;
        }

        HRESULT hr() const noexcept { return m_hr; }
        std::wstring source() const { return m_source; }
        std::wstring description() const { return m_description; }
        std::wstring help_file() const { return m_help_file; }
        DWORD help_context() const noexcept { return m_help_context; }
        GUID guid() const noexcept { return m_guid; }
    };

    inline std::wstring to_wstring(HRESULT const hr)
    {
        size_t const cch = 11;
        wchar_t buf[cch];
        swprintf(buf, cch, L"0x%08X", hr);
        return buf;
    }

    inline std::wstring to_wstring(REFGUID guid)
    {
        size_t const cch = 39;
        wchar_t buf[cch];
        swprintf(
            buf,
            cch,
            L"{%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
            guid.Data1,
            guid.Data2,
            guid.Data3,
            guid.Data4[0],
            guid.Data4[1],
            guid.Data4[2],
            guid.Data4[3],
            guid.Data4[4],
            guid.Data4[5],
            guid.Data4[6],
            guid.Data4[7]);
        return buf;
    }

    template<typename T>
    inline std::wstring to_wstring(T) = delete;
}

#endif  // COMEXCEPT_H

///////////////////////////////////////////////////////////////////////////////
