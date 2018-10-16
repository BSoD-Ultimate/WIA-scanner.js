#include "stdafx.h"
#include "WIADeviceMgr.h"

#include <experimental/filesystem>

namespace scanner
{
    // The callback used to write data fetched from WIA interface to file
    class CScanTransferCallback : public IWiaTransferCallback
    {
    public:
        CScanTransferCallback(CWIADevice& device,
            ATL::CComPtr<IWiaTransfer> transferInterface,
            const std::wstring& saveDirectory,
            const std::wstring& saveFilename,
            const std::wstring& fileExtension,
            bool isFeeder,
            ScanProgressCallback progressCallback = nullptr)
            : m_device(device)
            , m_pTransferInterface(transferInterface)
            , m_cRef(1)
            , m_fileIndex(0)
            , m_pageCount(0)
            , m_bFeeder(isFeeder)
            , m_saveDirectoryName(saveDirectory)
            , m_saveFilename(saveFilename)
            , m_fileExtension(fileExtension)
            , m_progressCallback(progressCallback)
        {
            assert(m_pTransferInterface);
            assert(!m_saveDirectoryName.empty());
            assert(!m_saveFilename.empty());
            assert(!m_fileExtension.empty());
        }
        virtual ~CScanTransferCallback()
        {
        }

        // IUnknown
        HRESULT CALLBACK QueryInterface(REFIID riid, void **ppvObject) override
        {
            if (NULL == ppvObject)
            {
                return E_INVALIDARG;
            }

            if (IsEqualIID(riid, IID_IUnknown))
            {
                *ppvObject = static_cast<IUnknown*>(this);
            }
            else if (IsEqualIID(riid, IID_IWiaTransferCallback))
            {
                *ppvObject = static_cast<IWiaTransferCallback*>(this);
            }
            else
            {
                *ppvObject = NULL;
                return (E_NOINTERFACE);
            }

            reinterpret_cast<IUnknown*>(*ppvObject)->AddRef();
            return S_OK;
        }
        ULONG CALLBACK AddRef() override
        {
            return InterlockedIncrement((long*)&m_cRef);
        }
        ULONG CALLBACK Release() override
        {
            LONG cRef = InterlockedDecrement((long*)&m_cRef);
            if (0 == cRef)
            {
                delete this;
            }
            return cRef;
        }

        // IWiaTransferCallback
        HRESULT STDMETHODCALLTYPE TransferCallback(LONG lFlags, WiaTransferParams* pWiaTransferParams) override
        {
            HRESULT hr = S_OK;

            if (pWiaTransferParams == NULL)
            {
                return E_INVALIDARG;
            }

            switch (pWiaTransferParams->lMessage)
            {
            case WIA_TRANSFER_MSG_STATUS:
            {
                if (!pWiaTransferParams->lPercentComplete)
                {
                    m_pageCount++;
                }

                ScanProgressInfo progressInfo;

                progressInfo.pageCount = m_pageCount;
                progressInfo.percentComplete = pWiaTransferParams->lPercentComplete;
                progressInfo.bytesTransferred = pWiaTransferParams->ulTransferredBytes;
                progressInfo.error = pWiaTransferParams->hrErrorStatus;

                if (m_progressCallback)
                {
                    m_progressCallback(progressInfo);
                }

                if (!m_device.IsScanRunning())
                {
                    HRESULT cancelResult = m_pTransferInterface->Cancel();
                    return E_ABORT;
                }
            }
            break;
            case WIA_TRANSFER_MSG_NEW_PAGE:
            {
            }
            break;
            case WIA_TRANSFER_MSG_DEVICE_STATUS:
            {
            }
            break;
            case WIA_TRANSFER_MSG_END_OF_STREAM:
            {
            }
            break;
            case WIA_TRANSFER_MSG_END_OF_TRANSFER:
            {
            }
            break;
            default:
                break;
            }

            return hr;
        }
        HRESULT STDMETHODCALLTYPE GetNextStream(LONG lFlags, BSTR bstrItemName, BSTR bstrFullItemName, IStream** ppDestination) override
        {
            HRESULT hr = S_OK;

            if ((!ppDestination) || (!bstrItemName) || (m_saveDirectoryName.empty()))
            {
                return E_INVALIDARG;
            }
            *ppDestination = NULL;

            const int pathMaxSize = 1000;
            std::unique_ptr<wchar_t[]> savePathBuf(new wchar_t[pathMaxSize]());

            if (!m_fileExtension.empty())
            {
                // Add page index after filename if the scanner is a feeder
                if (m_bFeeder)
                {
                    m_fileIndex++;
                    swprintf_s(savePathBuf.get(), pathMaxSize, L"%s\\%s_%d.%s", m_saveDirectoryName.c_str(), m_saveFilename.c_str(), m_fileIndex, m_fileExtension.c_str());
                }
                else
                {
                    swprintf_s(savePathBuf.get(), pathMaxSize, L"%s\\%s.%s", m_saveDirectoryName.c_str(), m_saveFilename.c_str(), m_fileExtension.c_str());
                }
            }
            else
            {
                swprintf_s(savePathBuf.get(), pathMaxSize, L"%s\\%s", m_saveDirectoryName.c_str(), m_saveFilename.c_str());
            }

            hr = SHCreateStreamOnFileW(savePathBuf.get(), STGM_CREATE | STGM_READWRITE, ppDestination);
            if (SUCCEEDED(hr))
            {
                m_saveFilePaths.push_back(savePathBuf.get());
            }
            else
            {
                *ppDestination = NULL;
            }
            return hr;
        }

        std::vector<std::wstring> GetSaveFilePaths() const
        {
            return m_saveFilePaths;
        }

    private:
        CWIADevice& m_device;
        ATL::CComPtr<IWiaTransfer> m_pTransferInterface;

        ULONG m_cRef;
        long m_fileIndex;
        bool m_bFeeder;                     // is the scanner a Feeder

        long m_pageCount;

        std::wstring m_saveDirectoryName;   // save directory
        std::wstring m_saveFilename;        // file name
        std::wstring m_fileExtension;       // file extension

        std::vector<std::wstring> m_saveFilePaths;

        ScanProgressCallback m_progressCallback;
    };

    CWIADeviceMgr::CWIADeviceMgr()
    {
        if (FAILED(CreateWIADeviveManager()) || FAILED(CreateWIADeviceInterfaceTable()))
        {
            throw std::runtime_error("Unable to create instance of IWiaDevMgr!");
        }
    }


    CWIADeviceMgr::~CWIADeviceMgr()
    {
    }

    ATL::CComPtr<IWiaDevMgr2> CWIADeviceMgr::get() const
    {
        return m_pWiaDevMgr;
    }

    std::unique_ptr<CWIADeviceMgr>& CWIADeviceMgr::GetInstance()
    {
        static std::unique_ptr<CWIADeviceMgr> instance(new CWIADeviceMgr());
        return instance;
    }

    ATL::CComPtr<IGlobalInterfaceTable> CWIADeviceMgr::GetInterfaceTable()
    {
        return m_pWiaDeviceTable;
    }

    std::shared_ptr<CWIADevice> CWIADeviceMgr::OpenWIADevice(const std::wstring& deviceId)
    {
        try
        {
            auto pDevice = std::make_shared<CWIADevice>(deviceId, *this);

            return pDevice;
        }
        catch (const std::exception&)
        {
            return nullptr;
        }

    }
    static void AssignDevicePropertyValue(WIADeviceProperties& props, const PROPSPEC& key, const PROPVARIANT& value)
    {
        if (key.ulKind == PRSPEC_PROPID)
        {
            if (key.propid == WIA_DIP_DEV_ID && value.vt == VT_BSTR)
            {
                props.deviceUUID = value.bstrVal;
            }
            if (key.propid == WIA_DIP_VEND_DESC && value.vt == VT_BSTR)
            {
                props.manufacturer = value.bstrVal;
            }
            if (key.propid == WIA_DIP_DEV_DESC && value.vt == VT_BSTR)
            {
                props.description = value.bstrVal;
            }
            if (key.propid == WIA_DIP_DEV_TYPE && value.vt == VT_I4)
            {
                props.type = value.intVal;
            }
            if (key.propid == WIA_DIP_PORT_NAME && value.vt == VT_BSTR)
            {
                props.port = value.bstrVal;
            }
            if (key.propid == WIA_DIP_DEV_NAME && value.vt == VT_BSTR)
            {
                props.deviceName = value.bstrVal;
            }
            if (key.propid == WIA_DIP_SERVER_NAME && value.vt == VT_BSTR)
            {
                props.server = value.bstrVal;
            }
            if (key.propid == WIA_DIP_REMOTE_DEV_ID && value.vt == VT_BSTR)
            {
                props.remoteDeviceId = value.bstrVal;
            }
            if (key.propid == WIA_DIP_UI_CLSID && value.vt == VT_BSTR)
            {
                props.uiClassId = value.bstrVal;
            }
            if (key.propid == WIA_DIP_HW_CONFIG && value.vt == VT_I4)
            {
                props.hardwareConfig = value.intVal;
            }
            if (key.propid == WIA_DIP_BAUDRATE && value.vt == VT_BSTR)
            {
                props.baudrate = value.bstrVal;
            }
            if (key.propid == WIA_DIP_STI_GEN_CAPABILITIES && value.vt == VT_I4)
            {
                props.STIGenericCapabilities = value.intVal;
            }
            if (key.propid == WIA_DIP_WIA_VERSION && value.vt == VT_BSTR)
            {
                props.WIAVersion = value.bstrVal;
            }
            if (key.propid == WIA_DIP_DRIVER_VERSION && value.vt == VT_BSTR)
            {
                props.DriverVersion = value.bstrVal;
            }
            if (key.propid == WIA_DIP_PNP_ID && value.vt == VT_BSTR)
            {
                props.PnPIDString = value.bstrVal;
            }
            if (key.propid == WIA_DIP_STI_DRIVER_VERSION && value.vt == VT_I4)
            {
                props.STIDriverVersion = value.intVal;
            }
        }
    }
    static std::shared_ptr<WIADeviceProperties> ReadDeviceProperties(ATL::CComPtr<IWiaPropertyStorage> deviceInfo)
    {
        HRESULT hr = S_OK;

        ULONG propertyValueCount = 0;
        hr = deviceInfo->GetCount(&propertyValueCount);

        // properties value list
        std::unique_ptr<PROPSPEC[]> propKeys(new PROPSPEC[propertyValueCount]());
        std::unique_ptr<PROPVARIANT[]> propValues(new PROPVARIANT[propertyValueCount]());

        // find what kind of property available for the current WIA device
        {
            ATL::CComPtr<IEnumSTATPROPSTG> propEnumrator;
            hr = deviceInfo->Enum(&propEnumrator);
            if (FAILED(hr))
            {
                return nullptr;
            }

            int propertyIndex = 0;
            while (hr == S_OK)
            {
                STATPROPSTG propertyInfo = { 0 };
                ULONG fetched = 0;
                hr = propEnumrator->Next(1, &propertyInfo, &fetched);

                propKeys[propertyIndex].ulKind = PRSPEC_PROPID;
                propKeys[propertyIndex].propid = propertyInfo.propid;

                if (propertyInfo.lpwstrName)
                {
                    CoTaskMemFree(propertyInfo.lpwstrName);
                }
                propertyIndex++;
            }
        }

        // Read property value
        // The last property value will be ignored. it is empty.
        hr = deviceInfo->ReadMultiple(propertyValueCount - 1, propKeys.get(), propValues.get());
        if (FAILED(hr))
        {
            return nullptr;
        }

        auto deviceProperties = std::make_shared<WIADeviceProperties>();

        for (size_t i = 0; i < propertyValueCount - 1; i++)
        {
            auto& propKey = propKeys[i];
            auto& propValue = propValues[i];

            AssignDevicePropertyValue(*deviceProperties, propKey, propValue);
        }
        FreePropVariantArray(propertyValueCount - 1, propValues.get());

        return deviceProperties;
    }

    std::vector<std::shared_ptr<WIADeviceProperties>> CWIADeviceMgr::ListAllDevices() const
    {
        std::vector<std::shared_ptr<WIADeviceProperties>> deviceInfo;
        assert(m_pWiaDevMgr);

        ATL::CComPtr<IEnumWIA_DEV_INFO> pWiaEnumDevInfo = NULL;
        HRESULT hr = m_pWiaDevMgr->EnumDeviceInfo(WIA_DEVINFO_ENUM_LOCAL, &pWiaEnumDevInfo);

        if (FAILED(hr))
        {
            return deviceInfo;
        }

        while (hr == S_OK)
        {
            ATL::CComPtr<IWiaPropertyStorage> pWiaPropertyStorage;
            hr = pWiaEnumDevInfo->Next(1, &pWiaPropertyStorage, NULL);

            if (hr == S_OK)
            {
                auto deviceProperty = ReadDeviceProperties(pWiaPropertyStorage);
                if (deviceProperty)
                {
                    deviceInfo.emplace_back(deviceProperty);
                }

            }
        }

        return deviceInfo;
    }

    HRESULT CWIADeviceMgr::CreateWIADeviveManager()
    {
        HRESULT hr = CoCreateInstance(CLSID_WiaDevMgr2, NULL, CLSCTX_LOCAL_SERVER, IID_IWiaDevMgr2, (void**)&m_pWiaDevMgr);
        return hr;
    }
    HRESULT CWIADeviceMgr::CreateWIADeviceInterfaceTable()
    {
        HRESULT hr = CoCreateInstance(CLSID_StdGlobalInterfaceTable, NULL, CLSCTX_INPROC_SERVER, IID_IGlobalInterfaceTable, (void**)&m_pWiaDeviceTable);
        return hr;
    }


    bool TraverseWIAItemTree(std::shared_ptr<WIAItemTreeNode> itemTree, WIAItemIterateCallback callback)
    {
        if (!callback)
        {
            return false;
        }

        bool ret = callback(itemTree->info);
        if (!ret)
        {
            return false;
        }
        for (size_t i = 0; i < itemTree->childItems.size(); i++)
        {
            bool childRet = TraverseWIAItemTree(itemTree->childItems[i], callback);
            if (!childRet)
            {
                return false;
            }
        }
        return true;
    }

    CWIADevice::CWIADevice(const std::wstring& deviceUUID, CWIADeviceMgr& manager)
        : m_manager(manager)
        , m_bScanRunning(false)
        , m_documentHandling(L"front")
        , m_imageFormat(L"tiff")
        , m_pageCount(ALL_PAGES)
    {

        ATL::CComPtr<IWiaItem2> pIWiaDevice;
        ATL::CComBSTR devId(deviceUUID.c_str());

        HRESULT hr = m_manager.get()->CreateDevice(0, devId, &pIWiaDevice);

        if (FAILED(hr))
        {
            throw std::runtime_error("unable to open wia device");
        }

        // Register the device in the IGlobalInterfaceTable
        m_pDevice = pIWiaDevice;
        auto interfaceTable = m_manager.GetInterfaceTable();
        assert(interfaceTable);

        hr = interfaceTable->RegisterInterfaceInGlobal(m_pDevice, IID_IWiaItem2, &m_deviceCookie);
        assert(SUCCEEDED(hr));

        m_imageSources = FindImageSourcesFromDevice(m_pDevice);
    }

    CWIADevice::~CWIADevice()
    {
        // Unregister the device from the IGlobalInterfaceTable
        auto interfaceTable = m_manager.GetInterfaceTable();
        HRESULT hr = interfaceTable->RevokeInterfaceFromGlobal(m_deviceCookie);
        assert(SUCCEEDED(hr));
    }

    ATL::CComPtr<IWiaItem2> CWIADevice::GetDevice()
    {
        ATL::CComPtr<IWiaItem2> device;

        auto interfaceTable = m_manager.GetInterfaceTable();
        HRESULT hr = interfaceTable->GetInterfaceFromGlobal(m_deviceCookie, IID_IWiaItem2, (void**)&device);
        return device;
    }

    std::vector<std::wstring> CWIADevice::GetImageSources()
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        std::vector<std::wstring> sources;

        TraverseWIAItemTree(m_imageSources,
            [&sources](const WIAItemTreeNodeInfo& info) -> bool
        {
            if (info.itemType & WiaItemTypeTransfer)
            {
                sources.push_back(info.deviceName);
            }
            return true;
        });

        return sources;
    }

    DWORD CWIADevice::GetInterfaceTableKey() const
    {
        return m_deviceCookie;
    }

    HRESULT CWIADevice::ShowDeviceDlg(HWND hWndParent, const std::wstring& folderName, const std::wstring& saveFilename, std::vector<std::wstring>& outFilePaths)
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        ATL::CComBSTR folderNameBstr(folderName.c_str());
        ATL::CComBSTR filenameBstr(saveFilename.c_str());

        // Create the directory specified first
        std::experimental::filesystem::create_directories(folderName);

        LONG fileCount = 0;
        BSTR* recvFilePaths = nullptr;
        IWiaItem2* recvFileObject = nullptr;
        HRESULT hr = m_pDevice->DeviceDlg(0, hWndParent, folderNameBstr, filenameBstr, &fileCount, &recvFilePaths, &recvFileObject);

        outFilePaths.clear();

        for (size_t i = 0; i < fileCount; i++)
        {
            outFilePaths.push_back(recvFilePaths[i]);
            SysFreeString(recvFilePaths[i]);
        }

        recvFileObject->Release();
        CoTaskMemFree(recvFilePaths);

        return hr;
    }

    std::wstring CWIADevice::GetImageFormat()
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);

        return m_imageFormat;
    }

    bool CWIADevice::SetImageFormat(const std::wstring & format)
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        if (!((format == L"tiff") ||
            (format == L"bmp") ||
            (format == L"jpeg") ||
            (format == L"png")))
        {
            return false;
        }

        m_imageFormat = format;
        return true;
    }

    std::wstring CWIADevice::GetColorFormat()
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        assert(m_imageSources);
        if (!m_imageSources)
        {
            return std::wstring();
        }

        auto& imageSource = GetImageSource(m_imageSources, L"");

        try
        {
            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = imageSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            LONG colorFormat = util::ReadPropertyLong(pIWiaPropertyStorage, WIA_IPA_DATATYPE);

            if (colorFormat == WIA_DATA_COLOR)
            {
                return L"fullcolor";
            }
            else if (colorFormat == WIA_DATA_GRAYSCALE)
            {
                return L"greyscale";
            }
            else if (colorFormat == WIA_DATA_THRESHOLD)
            {
                return L"blackwhite";
            }
            else if (colorFormat == WIA_DATA_AUTO)
            {
                return L"auto";
            }
            else
            {
                return std::wstring();
            }
        }
        catch (const util::PropertyStorageException&)
        {
            return std::wstring();
        }
    }

    bool CWIADevice::SetColorFormat(const std::wstring & format)
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        assert(m_imageSources);
        try
        {
            LONG colorFormat = WIA_DATA_THRESHOLD;
            if (format == L"blackwhite")
            {
                colorFormat = WIA_DATA_THRESHOLD;
            }
            else if (format == L"greyscale")
            {
                colorFormat = WIA_DATA_GRAYSCALE;
            }
            else if (format == L"fullcolor")
            {
                colorFormat = WIA_DATA_COLOR;
            }
            else
            {
                return false;
            }

            auto& imageSource = GetImageSource(m_imageSources, L"");

            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = imageSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            util::WritePropertyLong(pIWiaPropertyStorage, WIA_IPA_DATATYPE, colorFormat);
            return true;
        }
        catch (const util::PropertyStorageException&)
        {
            return false;
        }
    }
    struct PaperProfile
    {
        std::wstring paperProfileName;
        ULONG paperPropValue;
    };
    static const std::vector<PaperProfile> g_paperProfiles
    {
        { L"auto", WIA_PAGE_AUTO },
        { L"letter", WIA_PAGE_LETTER },
        { L"businesscard", WIA_PAGE_BUSINESSCARD },
        { L"uslegal", WIA_PAGE_USLEGAL },
        { L"usstatement", WIA_PAGE_USSTATEMENT },
        { L"a0", WIA_PAGE_ISO_A0 },
        { L"a1", WIA_PAGE_ISO_A1 },
        { L"a2", WIA_PAGE_ISO_A2 },
        { L"a3", WIA_PAGE_ISO_A3 },
        { L"a4", WIA_PAGE_ISO_A4 },
        { L"a5", WIA_PAGE_ISO_A5 },
        { L"a6", WIA_PAGE_ISO_A6 },
        { L"a7", WIA_PAGE_ISO_A7 },
        { L"a8", WIA_PAGE_ISO_A8 },
        { L"a9", WIA_PAGE_ISO_A9 },
        { L"a10", WIA_PAGE_ISO_A10 },
        { L"b0", WIA_PAGE_ISO_B0 },
        { L"b1", WIA_PAGE_ISO_B1 },
        { L"b2", WIA_PAGE_ISO_B2 },
        { L"b3", WIA_PAGE_ISO_B3 },
        { L"b4", WIA_PAGE_ISO_B4 },
        { L"b5", WIA_PAGE_ISO_B5 },
        { L"b6", WIA_PAGE_ISO_B6 },
        { L"b7", WIA_PAGE_ISO_B7 },
        { L"b8", WIA_PAGE_ISO_B8 },
        { L"b9", WIA_PAGE_ISO_B9 },
        { L"b10", WIA_PAGE_ISO_B10 },
    };

    std::wstring CWIADevice::GetPaperProfile()
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        assert(m_imageSources);
        if (!m_imageSources)
        {
            return L"";
        }

        auto& imageSource = GetImageSource(m_imageSources, L"");

        try
        {
            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = imageSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            LONG paperProfile = util::ReadPropertyLong(pIWiaPropertyStorage, WIA_IPS_PAGE_SIZE);

            auto paperProfileIter = std::find_if(g_paperProfiles.cbegin(), g_paperProfiles.cend(),
                [paperProfile](const PaperProfile& profile)
            {
                return profile.paperPropValue == paperProfile;
            });

            if (paperProfileIter == g_paperProfiles.cend())
            {
                return L"";
            }

            return paperProfileIter->paperProfileName;
        }
        catch (const util::PropertyStorageException&)
        {
            return L"";
        }
    }

    bool CWIADevice::SetPaperProfile(const std::wstring& profile)
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        assert(m_imageSources);
        try
        {
            auto paperProfileIter = std::find_if(g_paperProfiles.cbegin(), g_paperProfiles.cend(),
                [profile](const PaperProfile& profileData)
            {
                return profileData.paperProfileName == profile;
            });

            if (paperProfileIter == g_paperProfiles.cend())
            {
                return false;
            }

            auto& imageSource = GetImageSource(m_imageSources, L"");

            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = imageSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            util::WritePropertyLong(pIWiaPropertyStorage, WIA_IPS_PAGE_SIZE, paperProfileIter->paperPropValue);
            return true;
        }
        catch (const util::PropertyStorageException& e)
        {
            return false;
        }
    }

    std::wstring CWIADevice::GetDocumentHandling()
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);

        return m_documentHandling;
    }

    bool CWIADevice::SetDocumentHandling(const std::wstring & handling)
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        if (!((handling == L"front") ||
            (handling == L"duplex")))
        {
            return false;
        }

        m_documentHandling = handling;
        return true;
    }

    int CWIADevice::GetScanPageCount()
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        return m_pageCount;
    }

    bool CWIADevice::SetScanPageCount(int pageCount)
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        m_pageCount = pageCount;
        return true;
    }

    int CWIADevice::GetScanDPI()
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        assert(m_imageSources);
        if (!m_imageSources)
        {
            return 0;
        }

        auto& imageSource = GetImageSource(m_imageSources, L"");

        try
        {
            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = imageSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            LONG dpi = util::ReadPropertyLong(pIWiaPropertyStorage, WIA_IPS_XRES);

            return dpi;
        }
        catch (const util::PropertyStorageException&)
        {
            return 0;
        }

        return 0;
    }

    bool CWIADevice::SetScanDPI(int newDPI)
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        assert(m_imageSources);
        try
        {
            static const std::vector<int> commonDPIs{ 75, 100, 150, 200, 240, 250, 300, 400, 500, 600, 1200, 2400 };

            if (std::find(commonDPIs.cbegin(), commonDPIs.cend(), newDPI) == commonDPIs.cend())
            {
                return false;
            }

            auto& imageSource = GetImageSource(m_imageSources, L"");

            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = imageSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            util::WritePropertyLong(pIWiaPropertyStorage, WIA_IPS_XRES, newDPI);
            util::WritePropertyLong(pIWiaPropertyStorage, WIA_IPS_YRES, newDPI);
            return true;
        }
        catch (const util::PropertyStorageException&)
        {
            return false;
        }
    }

    bool CWIADevice::GetScanBrightnessRange(int& min, int& max, int& normal, int& step)
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        assert(m_imageSources);
        if (!m_imageSources)
        {
            return 0;
        }

        min = max = normal = step = 0;

        auto& imageSource = GetImageSource(m_imageSources, L"");
        try
        {
            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = imageSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            PROPSPEC brightnessSpec = { 0 };
            ULONG accessFlags = 0;
            util::CPropVariant props(1);

            brightnessSpec.ulKind = PRSPEC_PROPID;
            brightnessSpec.propid = WIA_IPS_BRIGHTNESS;

            hr = pIWiaPropertyStorage->GetPropertyAttributes(1, &brightnessSpec, &accessFlags, props.get());
            if (FAILED(hr))
            {
                throw util::PropertyStorageException("Error calling IWiaPropertyStorage::GetPropertyAttributes().", hr);
            }

            min = props[0].cal.pElems[WIA_RANGE_MIN];
            max = props[0].cal.pElems[WIA_RANGE_MAX];
            normal = props[0].cal.pElems[WIA_RANGE_NOM];
            step = props[0].cal.pElems[WIA_RANGE_STEP];

            return true;
        }
        catch (const util::PropertyStorageException&)
        {
            return false;
        }
    }

    bool CWIADevice::GetScanContrastRange(int & min, int & max, int& normal, int& step)
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        assert(m_imageSources);
        if (!m_imageSources)
        {
            return 0;
        }

        min = max = normal = step = 0;

        auto& imageSource = GetImageSource(m_imageSources, L"");
        try
        {
            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = imageSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            PROPSPEC brightnessSpec = { 0 };
            ULONG accessFlags = 0;
            util::CPropVariant props(1);

            brightnessSpec.ulKind = PRSPEC_PROPID;
            brightnessSpec.propid = WIA_IPS_CONTRAST;

            hr = pIWiaPropertyStorage->GetPropertyAttributes(1, &brightnessSpec, &accessFlags, props.get());
            if (FAILED(hr))
            {
                throw util::PropertyStorageException("Error calling IWiaPropertyStorage::GetPropertyAttributes().", hr);
            }

            min = props[0].cal.pElems[WIA_RANGE_MIN];
            max = props[0].cal.pElems[WIA_RANGE_MAX];
            normal = props[0].cal.pElems[WIA_RANGE_NOM];
            step = props[0].cal.pElems[WIA_RANGE_STEP];

            return true;
        }
        catch (const util::PropertyStorageException&)
        {
            return false;
        }
    }

    int CWIADevice::GetScanBrightness()
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        assert(m_imageSources);
        if (!m_imageSources)
        {
            return 0;
        }

        auto& imageSource = GetImageSource(m_imageSources, L"");

        try
        {
            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = imageSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            int min, max, normal, step;
            GetScanBrightnessRange(min, max, normal, step);

            LONG brightness = util::ReadPropertyLong(pIWiaPropertyStorage, WIA_IPS_BRIGHTNESS);

            return brightness / step;
        }
        catch (const util::PropertyStorageException&)
        {
            return 0;
        }

        return 0;
    }

    bool CWIADevice::SetScanBrightness(int brightness)
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        assert(m_imageSources);

        try
        {
            auto& imageSource = GetImageSource(m_imageSources, L"");

            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = imageSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            int min, max, normal, step;
            GetScanBrightnessRange(min, max, normal, step);

            util::WritePropertyLong(pIWiaPropertyStorage, WIA_IPS_BRIGHTNESS, brightness * step);
            return true;
        }
        catch (const util::PropertyStorageException&)
        {
            return false;
        }
    }

    int CWIADevice::GetScanContrast()
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        assert(m_imageSources);
        if (!m_imageSources)
        {
            return 0;
        }

        auto& imageSource = GetImageSource(m_imageSources, L"");

        try
        {
            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = imageSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            int min, max, normal, step;
            GetScanContrastRange(min, max, normal, step);

            LONG contrast = util::ReadPropertyLong(pIWiaPropertyStorage, WIA_IPS_CONTRAST);

            return contrast / step;
        }
        catch (const util::PropertyStorageException&)
        {
            return 0;
        }

        return 0;
    }

    bool CWIADevice::SetScanContrast(int contrast)
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);
        assert(m_imageSources);

        try
        {

            auto& imageSource = GetImageSource(m_imageSources, L"");

            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = imageSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            int min, max, normal, step;
            GetScanContrastRange(min, max, normal, step);

            util::WritePropertyLong(pIWiaPropertyStorage, WIA_IPS_BRIGHTNESS, contrast * step);
            return true;
        }
        catch (const util::PropertyStorageException&)
        {
            return false;
        }
    }

    bool CWIADevice::IsScanRunning() const
    {
        return m_bScanRunning;
    }

    void CWIADevice::CancelScan()
    {

        m_bScanRunning = false;
    }

    bool CWIADevice::IsFeeder()
    {
        HRESULT hr = S_OK;
        auto imgSource = GetImageSource(m_imageSources, L"");
        if (!imgSource)
        {
            return false;
        }

        ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
        hr = imgSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

        if (FAILED(hr))
        {
            return false;
        }

        GUID itemCategory = util::ReadPropertyGuid(pIWiaPropertyStorage, WIA_IPA_ITEM_CATEGORY);
        bool isFeeder = false;

        if (IsEqualGUID(itemCategory, WIA_CATEGORY_FEEDER))
        {
            isFeeder = true;
        }

        return isFeeder;
    }

    HRESULT CWIADevice::Scan(
        const std::wstring& saveDirectory,
        const std::wstring& saveFilename,
        std::vector<std::wstring>& saveFilePaths,
        ScanProgressCallback progressCallback)
    {
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);

        struct runningContext
        {
            runningContext(CWIADevice& d)
                : device(d)
            {
                device.m_bScanRunning = true;
            }
            ~runningContext()
            {
                device.m_bScanRunning = false;
            }
        private:
            CWIADevice& device;
        };
        runningContext c(*this);

        saveFilePaths.clear();

        // Get the first available image source pointer from the opened device.
        // We must acquire the device pointer from the IGlobalInterfaceTable object.
        // This method might be called from another thread apart from the thread where the object was created.
        // Windows will automatically marshal device pointer, then the error RPC_E_WRONG_THREAD will not occur.
        auto device = GetDevice();
        // Then rebuild the WIA item tree and find the first available device
        auto deviceTree = FindImageSourcesFromDevice(device);
        auto imgSource = GetImageSource(deviceTree, L"");
        if (!imgSource)
        {
            return E_FAIL;
        }

        // Query IWiaTransfer interface from the image source
        ATL::CComPtr<IWiaTransfer> pWiaTransfer;
        HRESULT hr = imgSource->QueryInterface(IID_IWiaTransfer, (void**)&pWiaTransfer);

        if (FAILED(hr))
        {
            return hr;
        }

        // Get image source property
        ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
        hr = imgSource->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

        if (FAILED(hr))
        {
            return hr;
        }
        try
        {
            // Check image source category here(is Feeder?)
            GUID itemCategory = util::ReadPropertyGuid(pIWiaPropertyStorage, WIA_IPA_ITEM_CATEGORY);
            bool isFeeder = false;

            // Set output file format
            SetDeviceImageFormat(imgSource, m_imageFormat);

            if (IsEqualGUID(itemCategory, WIA_CATEGORY_FEEDER))
            {
                // set feeder format
                SetDeviceDocumentHandling(imgSource, m_documentHandling);
                SetDeviceScanPageCount(imgSource, m_pageCount);
                isFeeder = true;
            }

            std::wstring fileExtension = m_imageFormat;
            if (!IsEqualIID(itemCategory, WIA_CATEGORY_FOLDER))
            {
                fileExtension = m_imageFormat;
            }

            // init callback
            ATL::CComPtr<IWiaTransferCallback> pCallback = new CScanTransferCallback(*this, pWiaTransfer, saveDirectory, saveFilename, fileExtension, isFeeder, progressCallback);

            hr = pWiaTransfer->Download(0, pCallback);

            saveFilePaths = ((CScanTransferCallback*)(&*pCallback))->GetSaveFilePaths();


        }
        catch (const util::PropertyStorageException& e)
        {
            return e.result;
        }

        return hr;
    }

    std::shared_ptr<WIAItemTreeNode> CWIADevice::FindImageSourcesFromDevice(ATL::CComPtr<IWiaItem2> pDevice)
    {
        return DoFindImageSourcesFromDevice(pDevice);
    }

    std::shared_ptr<WIAItemTreeNode> CWIADevice::DoFindImageSourcesFromDevice(ATL::CComPtr<IWiaItem2> pParentDevice)
    {
        assert(pParentDevice);

        // item type
        LONG lItemType = 0;
        HRESULT hr = pParentDevice->GetItemType(&lItemType);

        // query properties
        ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
        hr = pParentDevice->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

        std::wstring devName = util::ReadPropertyString(pIWiaPropertyStorage, WIA_IPA_FULL_ITEM_NAME);
        GUID itemCategory = util::ReadPropertyGuid(pIWiaPropertyStorage, WIA_IPA_ITEM_CATEGORY);

        auto pDeviceTreeNode = std::make_shared<WIAItemTreeNode>();
        pDeviceTreeNode->info.deviceName = devName;
        pDeviceTreeNode->info.itemType = lItemType;
        pDeviceTreeNode->info.itemCategory = itemCategory;
        pDeviceTreeNode->info.pWiaItem = pParentDevice;

        // folder
        // If it is a folder, enumerate its children
        if (lItemType & WiaItemTypeFolder)
        {
            // Get the child item enumerator for this item
            ATL::CComPtr<IEnumWiaItem2> pIEnumWiaItem2;
            hr = pParentDevice->EnumChildItems(0, &pIEnumWiaItem2);
            if (SUCCEEDED(hr))
            {
                // We will loop until we get an error or pEnumWiaItem->Next returns
                // S_FALSE to signal the end of the list.
                while (S_OK == hr)
                {
                    // Get the next child item
                    ATL::CComPtr<IWiaItem2> pChildWiaItem2;
                    hr = pIEnumWiaItem2->Next(1, &pChildWiaItem2, NULL);

                    // pEnumWiaItem->Next will return S_FALSE when the list is
                    // exhausted, so check for S_OK before using the returned
                    // value.
                    if (S_OK == hr)
                    {
                        // Recurse into this item
                        auto pChildItem = DoFindImageSourcesFromDevice(pChildWiaItem2);
                        pChildItem->pParent = pDeviceTreeNode;
                        pDeviceTreeNode->childItems.push_back(pChildItem);

                    }
                    else if (FAILED(hr))
                    {
                        // Report that an error occurred during enumeration
                        char errMsgBuf[200] = { 0 };
                        sprintf_s(errMsgBuf, "Error calling IEnumWiaItem2::Next. HRESULT=0x%08X", hr);
                        throw std::runtime_error(errMsgBuf);
                    }
                }

                // If the result of the enumeration is S_FALSE, since this
                // is normal, we will change it to S_OK
                if (S_FALSE == hr)
                {
                    hr = S_OK;
                }
            }
        }

        return pDeviceTreeNode;

    }
    ATL::CComPtr<IWiaItem2> CWIADevice::GetImageSource(std::shared_ptr<WIAItemTreeNode> itemTree, const std::wstring& imgSourceName)
    {
        assert(itemTree);
        if (!itemTree)
        {
            return nullptr;
        }
        if (!imgSourceName.empty())
        {
            ATL::CComPtr<IWiaItem2> ret;

            TraverseWIAItemTree(itemTree,
                [&ret, imgSourceName](const WIAItemTreeNodeInfo& info) -> bool
            {
                if (info.deviceName == imgSourceName)
                {
                    ret = info.pWiaItem;
                    return false;
                }
                return true;
            });

            return ret;
        }
        else
        {
            ATL::CComPtr<IWiaItem2> ret;

            TraverseWIAItemTree(itemTree,
                [&ret, imgSourceName](const WIAItemTreeNodeInfo& info) -> bool
            {
                if (info.itemType & WiaItemTypeTransfer)
                {
                    ret = info.pWiaItem;
                    return false;
                }
                return true;
            });

            return ret;
        }

    }
    bool CWIADevice::SetDeviceDocumentHandling(ATL::CComPtr<IWiaItem2> device, const std::wstring & handling)
    {
        assert(device);
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);

        try
        {
            LONG documentHandling = FRONT_ONLY;
            if (handling == L"front")
            {
                documentHandling = FRONT_ONLY;
            }
            else if (handling == L"duplex")
            {
                documentHandling = FEEDER | DUPLEX;
            }
            else
            {
                return false;
            }

            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = device->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            util::WritePropertyLong(pIWiaPropertyStorage, WIA_IPS_DOCUMENT_HANDLING_SELECT, documentHandling);

            return true;
        }
        catch (const util::PropertyStorageException& e)
        {
            return false;
        }
    }
    bool CWIADevice::SetDeviceImageFormat(ATL::CComPtr<IWiaItem2> device, const std::wstring & imageFormat)
    {
        assert(device);
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);

        try
        {
            GUID imageFormatGuid = GUID_NULL;
            LONG imageCompression = WIA_COMPRESSION_NONE;
            if (imageFormat == L"tiff")
            {
                imageFormatGuid = WiaImgFmt_TIFF;
            }
            else if (imageFormat == L"bmp")
            {
                imageFormatGuid = WiaImgFmt_BMP;
            }
            else if (imageFormat == L"jpeg")
            {
                imageFormatGuid = WiaImgFmt_JPEG;
                imageCompression = WIA_COMPRESSION_JPEG;
            }
            else if (imageFormat == L"png")
            {
                imageFormatGuid = WiaImgFmt_PNG;
                imageCompression = WIA_COMPRESSION_PNG;
            }
            else
            {
                return false;
            }

            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = device->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            util::WritePropertyLong(pIWiaPropertyStorage, WIA_IPA_COMPRESSION, imageCompression);
            util::WritePropertyGuid(pIWiaPropertyStorage, WIA_IPA_FORMAT, imageFormatGuid);


            return true;
        }
        catch (const util::PropertyStorageException& e)
        {
            return false;
        }
    }
    bool CWIADevice::SetDeviceScanPageCount(ATL::CComPtr<IWiaItem2> device, int pageCount)
    {
        assert(device);
        std::lock_guard<std::recursive_mutex> g(m_lockWIADevice);

        try
        {
            ATL::CComPtr<IWiaPropertyStorage> pIWiaPropertyStorage;
            HRESULT hr = device->QueryInterface(IID_IWiaPropertyStorage, (void**)&pIWiaPropertyStorage);

            util::WritePropertyLong(pIWiaPropertyStorage, WIA_IPS_PAGES, pageCount);

            return true;
        }
        catch (const util::PropertyStorageException& e)
        {
            return false;
        }
        return false;
    }

    DWORD CWIADevice::MarshalImgSource(ATL::CComPtr<IWiaItem2> device)
    {
        DWORD cookie = 0;
        auto interfaceTable = m_manager.GetInterfaceTable();
        assert(interfaceTable);

        HRESULT hr = interfaceTable->RegisterInterfaceInGlobal(m_pDevice, IID_IWiaItem2, &cookie);
        assert(SUCCEEDED(hr));
        return cookie;
    }
    ATL::CComPtr<IWiaItem2> CWIADevice::GetMarshalledImgSource(DWORD cookie)
    {
        ATL::CComPtr<IWiaItem2> device;

        auto interfaceTable = m_manager.GetInterfaceTable();
        HRESULT hr = interfaceTable->GetInterfaceFromGlobal(cookie, IID_IWiaItem2, (void**)&device);
        return device;
    }
    void CWIADevice::RemoveMarshalledImageSource(DWORD cookie)
    {
        auto interfaceTable = m_manager.GetInterfaceTable();
        HRESULT hr = interfaceTable->RevokeInterfaceFromGlobal(cookie);
        assert(SUCCEEDED(hr));
    }

}

