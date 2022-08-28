// test_ubstr.cpp: Test ComTools::UBSTR ///////////////////////////////////////
// Copyright (c) 2022, Jeffrey M. Engelmann

#include "CppUnitTest.h"
#include "ubstr.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;
using namespace ComTools;

///////////////////////////////////////////////////////////////////////////////

namespace TestComTools
{
    TEST_CLASS(TestUBSTR)
    {
        // Tests operator BSTR()
        void TakesBSTR(BSTR bstr)
        {
            Assert::AreEqual(L"This is a string.", bstr);
        }

        // Tests set()
        void GivesBSTR(BSTR* pbstr)
        {
            *pbstr = SysAllocString(L"This is a string.");
        }

    public:
        
        TEST_METHOD(InitDefault)
        {
            UBSTR s;                                            // tests default constructor
            Assert::AreEqual(false, (bool)s);                   // tests operator bool()
            Assert::AreEqual((size_t)0, s.length());            // tests length()
            Assert::AreEqual(std::wstring(), s.to_wstring());   // tests to_wstring()
        }

        TEST_METHOD(InitFromWsz)
        {
            const wchar_t wsz[] = L"This is a string.";
            UBSTR s(wsz);
            Assert::AreEqual(true, (bool)s);
            Assert::AreEqual(wcslen(wsz), s.length());
            Assert::AreEqual(std::wstring(wsz), s.to_wstring());
        }

        TEST_METHOD(InitFromWstring)
        {
            std::wstring ws(L"This is a string.");
            UBSTR s(ws);
            Assert::AreEqual(true, (bool)s);
            Assert::AreEqual(ws.length(), s.length());
            Assert::AreEqual(ws, s.to_wstring());
        }

        TEST_METHOD(TestTakesBSTR)
        {
            UBSTR s(L"This is a string.");
            TakesBSTR(s.get());
        }

        TEST_METHOD(TestSetBSTR)
        {
            UBSTR s;
            GivesBSTR(set(s));
            Assert::AreEqual(L"This is a string.", s.get());
        }

        TEST_METHOD(TestCopy)
        {
            UBSTR s1(L"This is a string.");
            UBSTR s2(s1);
            Assert::AreEqual(s1.to_wstring(), s2.to_wstring());
        }

        TEST_METHOD(TestCopyAssign)
        {
            UBSTR s1(L"This is a string.");
            UBSTR s2;
            s2 = s1;
            Assert::AreEqual(s1.to_wstring(), s2.to_wstring());
        }

        TEST_METHOD(TestMove)
        {
            UBSTR s1(L"This is a string.");
            UBSTR s2(std::move(s1));
            Assert::AreEqual(L"This is a string.", s2.to_wstring().c_str());
        }

        TEST_METHOD(TestMoveAssign)
        {
            UBSTR s1(L"This is a string.");
            UBSTR s2;
            s2 = std::move(s1);
            Assert::AreEqual(L"This is a string.", s2.to_wstring().c_str());
        }
    };
}

///////////////////////////////////////////////////////////////////////////////
