// test_comexcept.cpp: Test ComTools::ComExcept ///////////////////////////////
// Copyright (c) 2022, Jeffrey M. Engelmann

#include "CppUnitTest.h"
#include "comexcept.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;
using namespace ComTools;

///////////////////////////////////////////////////////////////////////////////

namespace TestCOMExcept
{
    HRESULT const demo_hr = E_FAIL;
    wchar_t const demo_source[] = L"SimulatedProgID.Object.1";
    wchar_t const demo_description[] = L"This is a simulated error message.\r\n";
    wchar_t const demo_help_file[] = L"C:\\Path\\To\\Simulated\\Help.chm";
    DWORD const demo_help_context = 2;
    GUID const demo_guid =
    { 0x3a8d763b, 0x8fc1, 0x40a0, { 0xb4, 0x95, 0xe1, 0x3d, 0xa6, 0x5f, 0x8b, 0x34 } };

    TEST_CLASS(TestComException)
    {
        static HRESULT FailWithoutErrorInfo()
        {
            return demo_hr;
        }

        static HRESULT FailWithBlankErrorInfo()
        {
            ICreateErrorInfo* pcei = nullptr;
            HRESULT hr = CreateErrorInfo(&pcei);
            if (SUCCEEDED(hr) && pcei)
            {
                IErrorInfo* pei = nullptr;
                hr = pcei->QueryInterface(IID_IErrorInfo, reinterpret_cast<void**>(&pei));
                if (SUCCEEDED(hr) && pei)
                {
                    hr = SetErrorInfo(0, pei);
                    pei->Release();
                }

                pcei->Release();
            }

            if (FAILED(hr)) throw std::exception("Could not set error info");
            return demo_hr;
        }

        static HRESULT FailWithErrorInfo()
        {
            ICreateErrorInfo* pcei = nullptr;
            HRESULT hr = CreateErrorInfo(&pcei);
            if (SUCCEEDED(hr) && pcei)
            {
                pcei->SetSource(const_cast<wchar_t*>(demo_source));
                pcei->SetDescription(const_cast<wchar_t*>(demo_description));
                pcei->SetHelpFile(const_cast<wchar_t*>(demo_help_file));
                pcei->SetHelpContext(demo_help_context);
                pcei->SetGUID(demo_guid);

                IErrorInfo* pei = nullptr;
                hr = pcei->QueryInterface(IID_IErrorInfo, reinterpret_cast<void**>(&pei));
                if (SUCCEEDED(hr) && pei)
                {
                    hr = SetErrorInfo(0, pei);
                    pei->Release();
                }

                pcei->Release();
            }

            if (FAILED(hr)) throw std::exception("Could not set error info");
            return demo_hr;
        }

        static void CheckErrorInfo(ComException const& e)
        {
            // Check the contents of e against the simulated error info
            Assert::AreEqual(demo_hr, e.hr());
            Assert::AreEqual(demo_source, e.source().c_str());
            Assert::AreEqual(demo_description, e.description().c_str());
            Assert::AreEqual(demo_help_file, e.help_file().c_str());
            Assert::AreEqual(demo_help_context, e.help_context());
            Assert::IsTrue(demo_guid == e.guid());
        }

        static void log(ComException& e)
        {
            // Log the contents of a COM exception object
            // This tests the custom to_wstring functions for HRESULT and GUID
            // that are implemented in comexcept.h
            Logger::WriteMessage((to_wstring(e.hr()) + L"\r\n").c_str());
            Logger::WriteMessage((e.source() + L"\r\n").c_str());
            Logger::WriteMessage(e.description().c_str());
            Logger::WriteMessage((e.help_file() + L"\r\n").c_str());
            Logger::WriteMessage((std::to_wstring(e.help_context()) + L"\r\n").c_str());
            Logger::WriteMessage((to_wstring(e.guid()) + L"\r\n").c_str());
        }

    public:

        TEST_METHOD(TestNoErrorInfo)
        {
            try
            {
                HRESULT hr = FailWithoutErrorInfo();
                if (FAILED(hr)) throw ComException(hr);
                Assert::Fail(L"Exception not thrown");
            }
            catch (ComException& e)
            {
                Assert::AreEqual(demo_hr, e.hr());
                Assert::IsTrue(e.source().empty());
                Assert::IsTrue(e.description().empty());
                Assert::IsTrue(e.help_file().empty());
                Assert::AreEqual(0UL, e.help_context());
                Assert::IsTrue(GUID_NULL == e.guid());

                log(e);
            }
            // Let other exceptions go uncaught to fail the test
        }

        TEST_METHOD(TestWithBlankErrorInfo)
        {
            try
            {
                HRESULT hr = FailWithBlankErrorInfo();
                if (FAILED(hr)) throw ComException(hr);
                Assert::Fail(L"Exception not thrown");
            }
            catch (ComException& e)
            {
                Assert::AreEqual(demo_hr, e.hr());
                Assert::IsTrue(e.source().empty());
                Assert::IsTrue(e.description().empty());
                Assert::IsTrue(e.help_file().empty());
                Assert::AreEqual(0UL, e.help_context());
                Assert::IsTrue(GUID_NULL == e.guid());

                log(e);
            }
            // Let other exceptions go uncaught to fail the test
        }

        TEST_METHOD(TestWithErrorInfo)
        {
            try
            {
                HRESULT hr = FailWithErrorInfo();
                if (FAILED(hr)) throw ComException(hr);
                Assert::Fail(L"Exception not thrown");
            }
            catch (ComException& e)
            {
                CheckErrorInfo(e);
                log(e);

                // Test the copy constructor
                ComException f(e);
                CheckErrorInfo(f);

                // Test the move constructor
                ComException g(std::move(f));
                CheckErrorInfo(g);

                // Test copy assignment
                ComException h;
                h = g;
                CheckErrorInfo(h);

                // Test move assignment
                ComException i;
                i = std::move(h);
                CheckErrorInfo(i);
            }
            // Let other exceptions go uncaught to fail the test
        }
    };
}

///////////////////////////////////////////////////////////////////////////////
