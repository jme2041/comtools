// ubstr.h ////////////////////////////////////////////////////////////////////
//
// ComTools::UBSTR C++ wrapper for BSTRs
//
// This is a shim class to wrap BSTRs that is based on class _UBSTR by Don Box.
// Box, D. (1998). Essential COM. Reading, MA: Addison-Wesley.
//
// ComTools::UBSTR is released under the MIT license.
//
// Copyright (c) 2022, Jeffrey M. Engelmann
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

#ifndef UBSTR_H
#define UBSTR_H

#include <Windows.h>
#include <utility>
#include <string>

namespace ComTools {
    class UBSTR {
        BSTR m_bstr = nullptr;

    public:
        friend void swap(UBSTR& a, UBSTR& b) noexcept
        {
            std::swap(a.m_bstr, b.m_bstr);
        }

        friend BSTR* set(UBSTR& obj) noexcept
        {
            if (obj.m_bstr)
            {
                SysFreeString(obj.m_bstr);
                obj.m_bstr = nullptr;
            }
            return &obj.m_bstr;
        }

        UBSTR() noexcept = default;

        ~UBSTR() noexcept { SysFreeString(m_bstr); }

        explicit UBSTR(wchar_t const* const wsz) noexcept :
            m_bstr(SysAllocString(wsz)) { }

        explicit UBSTR(std::wstring const& ws) noexcept :
            m_bstr(SysAllocString(ws.c_str())) { }

        UBSTR(UBSTR const& obj) noexcept :
            m_bstr(SysAllocString(obj.m_bstr)) { }

        UBSTR(UBSTR&& obj) noexcept : m_bstr() { swap(*this, obj); }

        UBSTR& operator=(UBSTR obj) noexcept
        {
            swap(*this, obj);
            return *this;
        }

        explicit operator bool() const noexcept { return m_bstr != nullptr; }

        size_t length() const noexcept
        {
            if (!*this) return 0;
            return wcslen(m_bstr);
        }

        std::wstring to_wstring() const
        {
            return *this ? std::wstring(m_bstr) : std::wstring();
        }

        BSTR get() const noexcept { return m_bstr; }
    };
}

#endif  // UBSTR_H

///////////////////////////////////////////////////////////////////////////////
