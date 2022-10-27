// test_ubstr.cpp: Test ComTools::IPtr ////////////////////////////////////////
// Copyright (c) 2022, Jeffrey M. Engelmann

#include "CppUnitTest.h"
#include <cstdio>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

#ifdef _DEBUG
#define IPTR_TRACE(s) Logger::WriteMessage(s"\r\n")
#endif

#include "iptr.h"

using namespace ComTools;

///////////////////////////////////////////////////////////////////////////////
//
// Simulate COM interfaces
//

#undef INTERFACE

#define INTERFACE IA
DECLARE_INTERFACE_IID_(IA, IUnknown, "B92F633A-8E96-11EB-B727-DC41A9695036")
{
    BEGIN_INTERFACE
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppv) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(Method1)(THIS_ wchar_t const* message) PURE;
    END_INTERFACE
};
#undef INTERFACE

#define INTERFACE IB
DECLARE_INTERFACE_IID_(IB, IUnknown, "B92F633B-8E96-11EB-B727-DC41A9695036")
{
    BEGIN_INTERFACE
    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void** ppv) PURE;
    STDMETHOD_(ULONG, AddRef)(THIS) PURE;
    STDMETHOD_(ULONG, Release)(THIS) PURE;
    STDMETHOD(Method2)(THIS_ IA * in) PURE;
    STDMETHOD(Method3)(THIS_ IA * *out) PURE;
    END_INTERFACE
};
#undef INTERFACE

class CAB : public IA, public IB {
    friend HRESULT NewCAB(REFIID, void**);

    ULONG m_rc;

    CAB() noexcept : m_rc(0)
    {
        size_t const cch = 64;
        char buf[cch];
        sprintf_s(buf, "(0x%016p): Creating CAB\r\n", this);
        Logger::WriteMessage(buf);
    }

public:
    virtual ~CAB() noexcept
    {
        size_t const cch = 64;
        char buf[cch];
        sprintf_s(buf, "(0x%016p): Destroying CAB\r\n", this);
        Logger::WriteMessage(buf);
    }

    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) noexcept override
    {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == __uuidof(IA)) *ppv = static_cast<IA*>(this);
        else if (riid == __uuidof(IB)) *ppv = static_cast<IB*>(this);
        else return (*ppv = nullptr), E_NOINTERFACE;
        reinterpret_cast<IUnknown*>(this)->AddRef();
        return S_OK;
    }

    STDMETHODIMP_(ULONG) AddRef() noexcept override
    {
        return ++m_rc;
    }

    STDMETHODIMP_(ULONG) Release() noexcept override
    {
        auto rc = --m_rc;
        if (rc == 0) delete this;
        return rc;
    }

    STDMETHODIMP Method1(wchar_t const* message) noexcept override
    {
        if (!message) return E_INVALIDARG;
        size_t const cch = 128;
        char buf[cch];
        sprintf_s(buf, "(0x%016p): IA::Method1: %S\r\n", this, message);
        Logger::WriteMessage(buf);
        return S_OK;
    }

    STDMETHODIMP Method2(IA* in) noexcept override
    {
        if (!in) return E_INVALIDARG;
        size_t const cch = 64;
        char buf[cch];
        sprintf_s(buf, "(0x%016p): IB::Method2\r\n", this);
        Logger::WriteMessage(buf);

        in->Method1(L"IB::Method2");
        return S_OK;
    }

    STDMETHODIMP Method3(IA** out) noexcept override
    {
        if (!out) return E_POINTER;
        *out = nullptr;

        size_t const cch = 64;
        char buf[cch];
        sprintf_s(buf, "(0x%016p): IB::Method3: Spawning a new CAB\r\n", this);
        Logger::WriteMessage(buf);

        try
        {
            auto p = new CAB();
            p->AddRef();
            HRESULT hr = p->QueryInterface(__uuidof(IA), reinterpret_cast<void**>(out));
            p->Release();
            return hr;
        }
        catch (std::bad_alloc&)
        {
            return E_OUTOFMEMORY;
        }
        catch (...)
        {
            return E_UNEXPECTED;
        }
    }
};

HRESULT NewCAB(REFIID riid, void** ppv)
{
    try
    {
        auto p = new CAB;
        p->AddRef();
        HRESULT hr = p->QueryInterface(riid, ppv);
        p->Release();
        return hr;
    }
    catch (std::bad_alloc&)
    {
        return E_OUTOFMEMORY;
    }
    catch (...)
    {
        return E_FAIL;
    }
}

namespace test_iptr
{
    TEST_CLASS(TestIPtr)
    {
        IPtr<IA> pA;
        IPtr<IB> pB;

    public:
        TEST_METHOD_INITIALIZE(AsGood)
        {
            // Create a CAB object
            HRESULT hr = NewCAB(__uuidof(IA), reinterpret_cast<void**>(set(pA)));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsTrue((bool)pA);

            // Test "As" with an interface that CAB implements
            pB = pA.As<IB>(__uuidof(IB));
            Assert::IsTrue((bool)pB);
        }

        TEST_METHOD(AsBad)
        {
            // Test "As" with an interface that CAB does not implement
            auto p = pA.As<ISupportErrorInfo>(__uuidof(ISupportErrorInfo));
            Assert::IsFalse((bool)p);
        }

        TEST_METHOD(QIGood)
        {
            // Test QueryInterface with an interface that CAB implements
            IPtr<IB> p;
            HRESULT hr = pA->QueryInterface(__uuidof(IB), reinterpret_cast<void**>(set(p)));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsTrue((bool)p);
        }

        TEST_METHOD(QIBad)
        {
            // Test QueryInterface with an interface that CAB does not implement
            IPtr<ISupportErrorInfo> p;
            HRESULT hr = pA->QueryInterface(__uuidof(ISupportErrorInfo), reinterpret_cast<void**>(set(p)));
            Assert::IsFalse(SUCCEEDED(hr));
            Assert::IsFalse((bool)p);
        }

        TEST_METHOD(Get)
        {
            // Call an IB method that takes an interface pointer
            HRESULT hr = pB->Method2(get(pA));
            Assert::IsTrue(SUCCEEDED(hr));
        }

        TEST_METHOD(Set)
        {
            // Call an IB method that returns an interface pointer
            IPtr<IA> p;
            HRESULT hr = pB->Method3(set(p));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsTrue((bool)p);
        }

        TEST_METHOD(Reset)
        {
            IPtr<IA> p;
            HRESULT hr = NewCAB(__uuidof(IA), reinterpret_cast<void**>(set(pA)));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsTrue((bool)pA);

            // Call set when p is not nullptr
            hr = pB->Method3(set(p));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsTrue((bool)p);
        }

        TEST_METHOD(Equality)
        {
            // Equality operator (pA and p point to the same object)
            auto p = pA.As<IA>(__uuidof(IA));
            Assert::IsTrue(pA == p);
        }

        TEST_METHOD(NonEquality)
        {
            // Not equals operator (pA and p point to different objects)
            IPtr<IA> p;
            HRESULT hr = NewCAB(__uuidof(IA), reinterpret_cast<void**>(set(p)));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsTrue(pA != p);
        }

        TEST_METHOD(Comparisons)
        {
            // Test other comparisons
            IPtr<IA> p;
            HRESULT hr = NewCAB(__uuidof(IA), reinterpret_cast<void**>(set(p)));
            Assert::IsTrue(SUCCEEDED(hr));

            if (pA < p)
            {
                Assert::IsTrue(p > pA);
                Assert::IsTrue(p >= pA);
                Assert::IsTrue(!(p <= pA));
            }
            else
            {
                Assert::IsTrue(pA > p);
                Assert::IsTrue(pA >= p);
                Assert::IsTrue(!(pA <= p));
            }
        }

        TEST_METHOD(CopyConstructSame)
        {
            // Copy construction (same interface type)
            auto p(pA);
            Assert::IsTrue((bool)p);
            p->Method1(L"I came from a copy constructor");
        }

        TEST_METHOD(CopyConstructCompatible)
        {
            // Template copy construction (different, but compatible, interface type)
            IPtr<IUnknown> p(pA);
            Assert::IsTrue((bool)p);
            p.As<IA>(__uuidof(IA))->Method1(L"I came from a different copy constructor");
        }

        TEST_METHOD(MoveConstruct)
        {
            // Move construction
            auto p1(pA);
            Assert::IsTrue((bool)p1);
            auto p2(std::move(p1));
            Assert::IsFalse((bool)p1);
            Assert::IsTrue((bool)p2);
            p2->Method1(L"I came from a move constructor");
        }

        TEST_METHOD(CopyAssignment)
        {
            // Copy assignment
            IPtr<IA> p;
            Assert::IsFalse((bool)p);
            p = pA;
            Assert::IsTrue((bool)p);
            p->Method1(L"I came from copy assignment");
        }

        TEST_METHOD(CopyAssignCompatible)
        {
            // Template copy assignment (different, but compatible, interface type)
            IPtr<IUnknown> p;
            p = pA;
            Assert::IsTrue((bool)p);
            p.As<IA>(__uuidof(IA))->Method1(L"I came from template copy assignment");
        }

        TEST_METHOD(MoveAssign)
        {
            auto p1(pA);
            Assert::IsTrue((bool)p1);
            IPtr<IA> p2;
            Assert::IsFalse((bool)p2);
            p2 = std::move(p1);
            Assert::IsFalse((bool)p1);
            Assert::IsTrue((bool)p2);
            p2->Method1(L"I came from move assignment");

        }

        TEST_METHOD(AssignNullptr)
        {
            auto p(pA);
            Assert::IsTrue((bool)p);
            p = nullptr;
            Assert::IsFalse((bool)p);
        }

        TEST_METHOD(DetachAttach)
        {
            auto p(pA);
            Assert::IsTrue((bool)p);

            // Detach
            IA* ia = detach(p);
            Assert::IsNotNull(ia);
            Assert::IsFalse((bool)p);

            // Attach
            attach(p, ia);
            ia = nullptr;
            Assert::IsTrue((bool)p);
            p->Method1(L"I came from attaching to a raw pointer");
        }

        TEST_METHOD(Swap)
        {
            auto p1(pA);
            Assert::IsTrue((bool)p1);
            IPtr<IA> p2;
            Assert::IsFalse((bool)p2);
            swap(p1, p2);
            Assert::IsFalse((bool)p1);
            Assert::IsTrue((bool)p2);
        }

        TEST_METHOD(CopyToNull)
        {
            auto p1(pA);
            Assert::IsTrue((bool)p1);
            Assert::AreEqual(E_POINTER, p1.CopyTo(nullptr));
        }

        TEST_METHOD(CopyTo)
        {
            // Hand out another reference to the same interface without QI
            auto p1(pA);
            Assert::IsTrue((bool)p1);
            IPtr<IA> p2;
            Assert::IsFalse((bool)p2);
            Assert::IsTrue(SUCCEEDED(p1.CopyTo(set(p2))));
            Assert::IsTrue((bool)p2);
            p2->Method1(L"I came from CopyTo");
        }

        TEST_METHOD(Release)
        {
            // Make a copy of an existing pointer that we would like to keep
            // We want this copy even after ia is released
            IA* ia = nullptr;
            HRESULT hr = pB->QueryInterface(__uuidof(IA), reinterpret_cast<void**>(&ia));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsNotNull(ia);
            IPtr<IA> p;
            p.CopyFrom(ia);
            ia->Release();
            ia = nullptr;
            Assert::IsTrue((bool)p);
            p->Method1(L"I came from CopyFrom");
        }
    };
}

///////////////////////////////////////////////////////////////////////////////
