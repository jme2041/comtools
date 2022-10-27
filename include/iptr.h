// iptr.h /////////////////////////////////////////////////////////////////////
//
// ComTools::IPtr: Smart pointer for COM interfaces
//
// IPtr is a non-Windows Runtime version of ComPtr, first implemented by Kenny
// Kerr and now incorporated into the Windows Runtime library (client.h).
// The original ComPtr was released under the MIT license.
// https://github.com/kennykerr/modern
//
// ComTools::IPtr is released under the MIT license.
//
// Copyright 2021-2022, Jeffrey M. Engelmann
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

#ifndef IPTR_H
#define IPTR_H

#include <Windows.h>

#ifndef IPTR_TRACE
#define IPTR_TRACE(s) ((void)0)
#endif

namespace ComTools {
    // This class hides AddRef() and Release()
    template<typename T>
    class NARR : public T {
        unsigned long __stdcall AddRef();
        unsigned long __stdcall Release();
    };

    // IPtr: Interface pointer wrapper. T is a COM interface.
    template<typename T>
    class IPtr {

        // U is a COM interface
        template<typename U>
        friend class IPtr;

        T* m_ptr = nullptr;

        void InternalAddRef() const noexcept
        {
            if (m_ptr) m_ptr->AddRef();
        }

        void InternalRelease() noexcept
        {
            T* temp = m_ptr;
            if (temp)
            {
                m_ptr = nullptr;
                temp->Release();
            }
        }

        void InternalCopy(T* other) noexcept
        {
            if (m_ptr != other)
            {
                InternalRelease();
                m_ptr = other;
                InternalAddRef();
            }
        }

        template<typename U>
        void InternalMove(IPtr<U>& other) noexcept
        {
            if (m_ptr != other.m_ptr)
            {
                InternalRelease();
                m_ptr = other.m_ptr;
                other.m_ptr = nullptr;
            }
        }

    public:
        IPtr() noexcept = default;

        IPtr(IPtr const& other) noexcept : m_ptr(other.m_ptr)
        {
            IPTR_TRACE("IPtr: Copy constructor");
            InternalAddRef();
        }

        template<typename U>
        explicit IPtr(IPtr<U> const& other) noexcept : m_ptr(other.m_ptr)
        {
            IPTR_TRACE("IPtr: Template copy constructor");
            InternalAddRef();
        }

        template<typename U>
        explicit IPtr(IPtr<U>&& other) noexcept : m_ptr(other.m_ptr)
        {
            IPTR_TRACE("IPtr: Move constructor");
            other.m_ptr = nullptr;
        }

        ~IPtr() noexcept
        {
            IPTR_TRACE("IPtr: Destructor");
            InternalRelease();
        }

        IPtr& operator=(IPtr const& other) noexcept
        {
            IPTR_TRACE("IPtr: Copy assignment");
            InternalCopy(other.m_ptr);
            return *this;
        }

        template<typename U>
        IPtr& operator=(IPtr<U> const& other) noexcept
        {
            IPTR_TRACE("IPtr: Template copy assignment");
            InternalCopy(other.m_ptr);
            return *this;
        }

        template<typename U>
        IPtr& operator=(IPtr<U>&& other) noexcept
        {
            IPTR_TRACE("IPtr: Move assignment");
            InternalMove(other);
            return *this;
        }

        IPtr& operator=(std::nullptr_t) noexcept
        {
            IPTR_TRACE("IPtr: nullptr assignment");
            InternalRelease();
            return *this;
        }

        explicit operator bool() const noexcept
        {
            return m_ptr != nullptr;
        }

        NARR<T>* operator->() const noexcept
        {
            return static_cast<NARR<T>*>(m_ptr);
        }

        friend T* get(IPtr const& obj) noexcept
        {
            return obj.m_ptr;
        }

        friend T** set(IPtr& obj) noexcept
        {
            if (obj) obj = nullptr;
            return &obj.m_ptr;
        }

        friend void attach(IPtr& obj, T* p) noexcept
        {
            obj.InternalRelease();
            obj.m_ptr = p;
        }

        friend T* detach(IPtr& obj) noexcept
        {
            T* temp = obj.m_ptr;
            obj.m_ptr = nullptr;
            return temp;
        }

        friend void swap(IPtr& left, IPtr& right) noexcept
        {
            T* temp = left.m_ptr;
            left.m_ptr = right.m_ptr;
            right.m_ptr = temp;
        }

        template<typename U>
        IPtr<U> As(REFIID riid) const noexcept
        {
            IPtr<U> temp;
            m_ptr->QueryInterface(
                riid,
                reinterpret_cast<void**>(set(temp)));
            return temp;
        }

        void CopyFrom(T* other) noexcept
        {
            InternalCopy(other);
        }

        HRESULT CopyTo(T** other) noexcept
        {
            if (!other) return E_POINTER;
            InternalAddRef();
            *other = m_ptr;
            return S_OK;
        }
    };

    template<typename T>
    bool operator==(IPtr<T> const& left, IPtr<T> const& right) noexcept
    {
        return get(left) == get(right);
    }

    template<typename T>
    bool operator!=(IPtr<T> const& left, IPtr<T> const& right) noexcept
    {
        return !(left == right);
    }

    template<typename T>
    bool operator<(IPtr<T> const& left, IPtr<T> const& right) noexcept
    {
        return get(left) < get(right);
    }

    template<typename T>
    bool operator>(IPtr<T> const& left, IPtr<T> const& right) noexcept
    {
        return right < left;
    }

    template<typename T>
    bool operator<=(IPtr<T> const& left, IPtr<T> const& right) noexcept
    {
        return !(right < left);
    }

    template<typename T>
    bool operator>=(IPtr<T> const& left, IPtr<T> const& right) noexcept
    {
        return !(left < right);
    }
}

#endif // IPTR_H

///////////////////////////////////////////////////////////////////////////////
