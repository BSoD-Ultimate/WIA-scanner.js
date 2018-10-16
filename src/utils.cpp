#include "stdafx.h"
#include "utils.h"
#include <comdef.h>

namespace scanner
{
    namespace util
    {
        COMEnvironment::COMEnvironment(DWORD mode)
        {
            CoInitializeEx(NULL, mode);
        }
        COMEnvironment::~COMEnvironment()
        {
            CoUninitialize();
        }

        DWORD PointerToIdValue(void* ptr)
        {
#if defined(__i386) || defined(__i386__) || defined(_M_IX86)
            DWORD id = (DWORD)ptr;
#elif defined(__x86_64) || defined(__x86_64__) || defined(__amd64) || defined(_M_X64)
            uintptr_t ptrValue = uintptr_t(ptr);
            DWORD id = (ptrValue & 0xffffffff) ^ ((ptrValue >> 32) & 0xffffffff);
#endif
            return id;
        }

        std::string WStringToUTF8(const std::wstring & wstr)
        {
            int nBufSizeUTF8 = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), wstr.length(), NULL, 0, NULL, NULL);
            std::unique_ptr<char[]> pBufUTF8(new CHAR[nBufSizeUTF8]());
            WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), wstr.length(), pBufUTF8.get(), nBufSizeUTF8, NULL, NULL);
            std::string result(pBufUTF8.get(), nBufSizeUTF8);
            return result;
        }

        std::wstring WStringFromUTF8(const std::string & utf8Str)
        {
            int nLengthResult = MultiByteToWideChar(CP_UTF8, 0, utf8Str.c_str(), utf8Str.length(), NULL, 0);
            std::unique_ptr<wchar_t[]> bufResult(new wchar_t[nLengthResult]());
            MultiByteToWideChar(CP_UTF8, 0, utf8Str.c_str(), utf8Str.length(), bufResult.get(), nLengthResult);
            std::wstring strResult(bufResult.get(), nLengthResult);
            return strResult;
        }

        CPropVariant::CPropVariant(int propCount)
            : m_pPropValues(new PROPVARIANT[propCount]())
        {
            PropVariantInit(m_pPropValues.get());
        }
        CPropVariant::~CPropVariant()
        {
            PropVariantClear(m_pPropValues.get());
        }
        const PROPVARIANT& CPropVariant::operator[](int index) const
        {
            return m_pPropValues[index];
        }
        PROPVARIANT & CPropVariant::operator[](int index)
        {
            return m_pPropValues[index];
        }

        PROPVARIANT* CPropVariant::get() const
        {
            return m_pPropValues.get();
        }

        // WIA properties
        // This function reads item property which returns BSTR like WIA_DIP_DEV_ID, WIA_DIP_DEV_NAME etc.
        std::wstring ReadPropertyString(ATL::CComPtr<IWiaPropertyStorage> pWiaPropertyStorage, PROPID propid)
        {
            assert(pWiaPropertyStorage);
            if ((!pWiaPropertyStorage))
            {
                throw PropertyStorageException("invalid IWiaPropertyStorage pointer", E_INVALIDARG);
            }
            //initialize out variables
            std::wstring strValue;

            // Declare PROPSPECs and PROPVARIANTs, and initialize them.
            PROPSPEC PropSpec[1] = { 0 };
            CPropVariant PropVar(1);

            PropSpec[0].ulKind = PRSPEC_PROPID;
            PropSpec[0].propid = propid;

            HRESULT hr = pWiaPropertyStorage->ReadMultiple(1, PropSpec, PropVar.get());
            if (S_OK == hr)
            {
                if (PropVar[0].vt == VT_BSTR)
                {
                    strValue = PropVar[0].bstrVal;
                }
                else
                {
                    throw PropertyStorageException("trying to read a property which type is not a BSTR", E_INVALIDARG);
                }
            }
            else
            {
                throw PropertyStorageException("Error calling IWiaPropertyStorage::ReadMultiple().", hr);
            }

            return strValue;
        }

        // This function reads item property which returns LONG 
        LONG ReadPropertyLong(ATL::CComPtr<IWiaPropertyStorage> pWiaPropertyStorage, PROPID propid)
        {
            assert(pWiaPropertyStorage);
            if ((!pWiaPropertyStorage))
            {
                throw PropertyStorageException("invalid IWiaPropertyStorage pointer", E_INVALIDARG);
            }
            LONG value = 0;

            // Declare PROPSPECs and PROPVARIANTs, and initialize them.
            PROPSPEC PropSpec[1] = { 0 };
            CPropVariant PropVar(1);

            PropSpec[0].ulKind = PRSPEC_PROPID;
            PropSpec[0].propid = propid;
            HRESULT hr = pWiaPropertyStorage->ReadMultiple(1, PropSpec, PropVar.get());
            if (S_OK == hr)
            {
                if (PropVar[0].vt == VT_I4)
                {
                    value = PropVar[0].lVal;
                }
                else
                {
                    throw PropertyStorageException("trying to read a property which type is not a LONG", E_INVALIDARG);
                }
            }
            else
            {
                throw PropertyStorageException( "Error calling IWiaPropertyStorage::ReadMultiple().", hr);
            }

            return value;
        }

        // This function reads item property which returns GUID like WIA_IPA_FORMAT, WIA_IPA_ITEM_CATEGORY etc.
        GUID ReadPropertyGuid(ATL::CComPtr<IWiaPropertyStorage> pWiaPropertyStorage, PROPID propid)
        {
            assert(pWiaPropertyStorage);
            if ((!pWiaPropertyStorage))
            {
                throw PropertyStorageException("invalid IWiaPropertyStorage pointer", E_INVALIDARG);
            }
            //initialize out variables
            GUID guid_val = GUID_NULL;

                // Declare PROPSPECs and PROPVARIANTs, and initialize them.
                PROPSPEC PropSpec[1] = { 0 };
                CPropVariant PropVar(1);

                PropSpec[0].ulKind = PRSPEC_PROPID;
                PropSpec[0].propid = propid;
                HRESULT hr = pWiaPropertyStorage->ReadMultiple(1, PropSpec, PropVar.get());
                if (S_OK == hr)
                {
                    if (PropVar[0].vt == VT_CLSID)
                    {
                        memcpy(&guid_val, PropVar[0].puuid, sizeof(GUID));
                    }
                    else
                    {
                        throw PropertyStorageException("trying to read a property which type is not a GUID", E_INVALIDARG);
                    }

                }
                else
                {
                    throw PropertyStorageException( "Error calling IWiaPropertyStorage::ReadMultiple().", hr);
                }

            return guid_val;
        }

        // This function writes item property which takes LONG like WIA_IPA_PAGES etc.
        void WritePropertyString(ATL::CComPtr<IWiaPropertyStorage> pWiaPropertyStorage, PROPID propid, const std::wstring& value)
        {
            assert(pWiaPropertyStorage);
            if ((!pWiaPropertyStorage))
            {
                throw PropertyStorageException("invalid IWiaPropertyStorage pointer", E_INVALIDARG);
            }

            // Declare PROPSPECs and PROPVARIANTs, and initialize them.
            PROPSPEC PropSpec[1] = { 0 };
            PROPVARIANT PropVar[1];
            PropVariantInit(PropVar);

            ATL::CComBSTR valueBstr(value.c_str());

            //Fill values in Propvar which are to be written 
            PropSpec[0].ulKind = PRSPEC_PROPID;
            PropSpec[0].propid = propid;
            PropVar[0].vt = VT_BSTR;
            PropVar[0].bstrVal = valueBstr;

            HRESULT hr = pWiaPropertyStorage->WriteMultiple(1, PropSpec, PropVar, WIA_DIP_FIRST);

            if (FAILED(hr))
            {
                throw PropertyStorageException("Error calling IWiaPropertyStorage::WriteMultiple().", hr);
            }
        }

        // This function writes item property which takes LONG like WIA_IPA_PAGES etc.
        void WritePropertyLong(ATL::CComPtr<IWiaPropertyStorage> pWiaPropertyStorage, PROPID propid, LONG lVal)
        {
            assert(pWiaPropertyStorage);
            if ((!pWiaPropertyStorage))
            {
                throw PropertyStorageException("invalid IWiaPropertyStorage pointer", E_INVALIDARG);
            }

            // Declare PROPSPECs and PROPVARIANTs, and initialize them.
            PROPSPEC PropSpec[1] = { 0 };
            PROPVARIANT PropVar[1];
            PropVariantInit(PropVar);

            //Fill values in Propvar which are to be written 
            PropSpec[0].ulKind = PRSPEC_PROPID;
            PropSpec[0].propid = propid;
            PropVar[0].vt = VT_I4;
            PropVar[0].lVal = lVal;

            HRESULT hr = pWiaPropertyStorage->WriteMultiple(1, PropSpec, PropVar, WIA_DIP_FIRST);

            if (FAILED(hr))
            {
                throw PropertyStorageException("Error calling IWiaPropertyStorage::WriteMultiple().", hr);
            }
        }

        // This function writes item property which takes GUID like WIA_IPA_FORMAT, WIA_IPA_ITEM_CATEGORY etc.
        void WritePropertyGuid(ATL::CComPtr<IWiaPropertyStorage> pWiaPropertyStorage, PROPID propid, GUID guid)
        {
            assert(pWiaPropertyStorage);
            if ((!pWiaPropertyStorage))
            {
                throw PropertyStorageException("invalid IWiaPropertyStorage pointer", E_INVALIDARG);
            }

            // Declare PROPSPECs and PROPVARIANTs, and initialize them.
            PROPSPEC PropSpec[1] = { 0 };
            PROPVARIANT PropVar[1];
            PropVariantInit(PropVar);

            PropSpec[0].ulKind = PRSPEC_PROPID;
            PropSpec[0].propid = propid;

            //Fill values in Propvar which are to be written 
            PropVar[0].vt = VT_CLSID;
            PropVar[0].puuid = &guid;

            HRESULT hr = pWiaPropertyStorage->WriteMultiple(1, PropSpec, PropVar, WIA_DIP_FIRST);

            if (FAILED(hr))
            {
                throw PropertyStorageException("Error calling IWiaPropertyStorage::WriteMultiple().", hr);
            }
        }

        std::wstring GetWIAErrorStr(HRESULT ret)
        {
            _com_error err(ret);
            std::wstring errorMsg = err.ErrorMessage();

            if (errorMsg.find(L"Unknown error") != std::wstring::npos)
            {
                switch (ret)
                {
                case WIA_ERROR_GENERAL_ERROR:
                    errorMsg = L"Unknown error.";
                    break;
                case WIA_ERROR_PAPER_JAM:
                    errorMsg = L"Paper jam.";
                    break;
                case WIA_ERROR_PAPER_EMPTY:
                    errorMsg = L"No paper detected.";
                    break;
                case WIA_ERROR_PAPER_PROBLEM:
                    errorMsg = L"An error has occurred on the feeder.";
                    break;
                case WIA_ERROR_OFFLINE:
                    errorMsg = L"Device is offline.";
                    break;
                case WIA_ERROR_BUSY:
                    errorMsg = L"Device is busy.";
                    break;
                case WIA_ERROR_WARMING_UP:
                    errorMsg = L"Device is warming up.";
                    break;
                case WIA_ERROR_USER_INTERVENTION:
                    errorMsg = L"User error, please check if the device is correctly connected.";
                    break;
                case WIA_ERROR_ITEM_DELETED:
                    errorMsg = L"Device is deleted.";
                    break;
                case WIA_ERROR_DEVICE_COMMUNICATION:
                    errorMsg = L"Error connecting to the device.";
                    break;
                case WIA_ERROR_INVALID_COMMAND:
                    errorMsg = L"Invalid device command.";
                    break;
                case WIA_ERROR_INCORRECT_HARDWARE_SETTING:
                    errorMsg = L"Incorrect hardware setting.";
                    break;
                case WIA_ERROR_DEVICE_LOCKED:
                    errorMsg = L"Device is locked by other applications.";
                    break;
                case WIA_ERROR_EXCEPTION_IN_DRIVER:
                    errorMsg = L"An exception has occurred in the driver";
                    break;
                case WIA_ERROR_INVALID_DRIVER_RESPONSE:
                    errorMsg = L"Invalid response from the driver.";
                    break;
                case WIA_ERROR_COVER_OPEN:
                    errorMsg = L"Cover of the scanner is opened.";
                    break;
                case WIA_ERROR_LAMP_OFF:
                    errorMsg = L"Light bulb malfunctioning.";
                    break;
                case WIA_ERROR_DESTINATION:
                    errorMsg = L"Unknown error.";
                    break;
                case WIA_ERROR_MULTI_FEED:
                    errorMsg = L"Multi papers was feeded into the feeder.";
                    break;
                default:
                    break;
                }
            }

            return errorMsg;
        }

    }
}
