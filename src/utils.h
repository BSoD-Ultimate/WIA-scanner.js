#pragma once

#include <wia.h>

namespace scanner
{
    namespace util
    {

        // Manages COM environment in a thread
        class COMEnvironment
        {
        public:
            COMEnvironment(DWORD mode = COINIT_MULTITHREADED);
            ~COMEnvironment();

            COMEnvironment(const COMEnvironment&) = delete;
            COMEnvironment& operator=(const COMEnvironment&) = delete;

        };

        DWORD PointerToIdValue(void* ptr);

        class PropertyStorageException : public std::runtime_error
        {
        public:
            PropertyStorageException(const std::string& msg, HRESULT result)
                : std::runtime_error(msg)
                , result(result)
            {
            }

            HRESULT result;
        };

        std::string WStringToUTF8(const std::wstring& wstr);
        std::wstring WStringFromUTF8(const std::string& utf8Str);

        std::wstring ReadPropertyString(ATL::CComPtr<IWiaPropertyStorage> pWiaPropertyStorage, PROPID propid);

        LONG ReadPropertyLong(ATL::CComPtr<IWiaPropertyStorage> pWiaPropertyStorage, PROPID propid);

        GUID ReadPropertyGuid(ATL::CComPtr<IWiaPropertyStorage> pWiaPropertyStorage, PROPID propid);
        void WritePropertyString(ATL::CComPtr<IWiaPropertyStorage> pWiaPropertyStorage, PROPID propid, const std::wstring & value);
        void WritePropertyLong(ATL::CComPtr<IWiaPropertyStorage> pWiaPropertyStorage, PROPID propid, LONG lVal);
        void WritePropertyGuid(ATL::CComPtr<IWiaPropertyStorage> pWiaPropertyStorage, PROPID propid, GUID guid);

        class CPropVariant
        {
        public:
            CPropVariant(int propCount);
            ~CPropVariant();

            CPropVariant(const CPropVariant&) = delete;
            CPropVariant& operator=(const CPropVariant&) = delete;

            const PROPVARIANT& operator[](int index) const;
            PROPVARIANT& operator[](int index);

            PROPVARIANT* get() const;

        private:
            std::unique_ptr<PROPVARIANT[]> m_pPropValues;

        };

        // get error message string from the WIA error code
        std::wstring GetWIAErrorStr(HRESULT ret);
    }
}
