/*
 * Wrapper of the WIA(Windows Image Acquisition) API
*/
#pragma once

#include <wia.h>

#include <mutex>
#include <map>

namespace scanner
{
    class CWIADevice;

    struct WIADeviceProperties
    {
        std::wstring deviceUUID;
        std::wstring manufacturer;
        std::wstring description;
        int32_t type;
        std::wstring port;
        std::wstring deviceName;
        std::wstring server;
        std::wstring remoteDeviceId;
        std::wstring uiClassId;
        int32_t hardwareConfig;
        std::wstring baudrate;
        int32_t STIGenericCapabilities;
        std::wstring WIAVersion;
        std::wstring DriverVersion;
        std::wstring PnPIDString;
        int32_t STIDriverVersion;

        WIADeviceProperties()
            : type(0), hardwareConfig(0), STIGenericCapabilities(0), STIDriverVersion(0)
        {
        }
    };

    class CWIADeviceMgr
    {
    private:
        CWIADeviceMgr();
    public:
        ~CWIADeviceMgr();
        ATL::CComPtr<IWiaDevMgr2> get() const;

        static std::unique_ptr<CWIADeviceMgr>& GetInstance();

        ATL::CComPtr<IGlobalInterfaceTable> GetInterfaceTable();

        std::shared_ptr<CWIADevice> OpenWIADevice(const std::wstring& deviceId);
        std::vector<std::shared_ptr<WIADeviceProperties>> ListAllDevices() const;

    private:
        HRESULT CreateWIADeviveManager();
        HRESULT CreateWIADeviceInterfaceTable();

    private:
        ATL::CComPtr<IWiaDevMgr2> m_pWiaDevMgr;
        ATL::CComPtr<IGlobalInterfaceTable> m_pWiaDeviceTable;
    };

    // information of the scan progress
    struct ScanProgressInfo
    {
        LONG pageCount;
        LONG percentComplete;
        ULONG64 bytesTransferred;
        HRESULT error;

        ScanProgressInfo()
            : pageCount(0)
            , percentComplete(0)
            , bytesTransferred(0)
            , error(S_OK)
        {
        }
    };
    typedef std::function<void(const ScanProgressInfo&)> ScanProgressCallback;

    struct WIAItemTreeNodeInfo
    {
        std::wstring deviceName;
        ATL::CComPtr<IWiaItem2> pWiaItem;
        LONG itemType = 0;
        GUID itemCategory = { 0 };
    };

    struct WIAItemTreeNode
    {
        std::weak_ptr<WIAItemTreeNode> pParent;
        WIAItemTreeNodeInfo info;
        std::vector<std::shared_ptr<WIAItemTreeNode>> childItems;
    };

    // Traverse a WIA item tree
    typedef std::function<bool(WIAItemTreeNodeInfo&)> WIAItemIterateCallback;
    bool TraverseWIAItemTree(std::shared_ptr<WIAItemTreeNode> itemTree, WIAItemIterateCallback callback);


    class CWIADevice
    {
    public:
        CWIADevice(const std::wstring& deviceUUID, CWIADeviceMgr& manager);
        ~CWIADevice();

        CWIADevice(const CWIADevice&) = delete;
        CWIADevice& operator=(const CWIADevice&) = delete;

        // IWiaItem object pointer must be get from a IGlobalInterfaceTable object under multi-threaded circumstances.
        // Otherwise, the error RPC_E_WRONG_THREAD will occur
        ATL::CComPtr<IWiaItem2> GetDevice();
        std::vector<std::wstring> GetImageSources();

        DWORD GetInterfaceTableKey() const;

        HRESULT ShowDeviceDlg(HWND hWndParent, const std::wstring& folderName, const std::wstring& saveFilename, std::vector<std::wstring>& outFilePaths);

        // Read/write properties of the WIA device
        // image output format(tiff/bmp/jpeg/png)
        std::wstring GetImageFormat();
        bool SetImageFormat(const std::wstring& format);
        // color mode(blackwhite/greyscale/fullcolor)
        std::wstring GetColorFormat();
        bool SetColorFormat(const std::wstring& format);
        // paper profile
        std::wstring GetPaperProfile();
        bool SetPaperProfile(const std::wstring& profile);
        // document handling(front/duplex). Available only if the scanner is a feeder
        std::wstring GetDocumentHandling();
        bool SetDocumentHandling(const std::wstring& handling);
        // page count(Available only if the scanner is a feeder, will keep scanning until no more papers if the value is 0.)
        int GetScanPageCount();
        bool SetScanPageCount(int pageCount);
        // DPI(75/100/150/200/240/250/300/400/500/600/1200/2400)
        int GetScanDPI();
        bool SetScanDPI(int newDPI);
        // brightness/contrast range
        bool GetScanBrightnessRange(int& min, int& max, int& normal, int& step);
        bool GetScanContrastRange(int& min, int& max, int& normal, int& step);
        // brightness
        int GetScanBrightness();
        bool SetScanBrightness(int brightness);
        // contrast
        int GetScanContrast();
        bool SetScanContrast(int contrast);

        bool IsScanRunning() const;
        void CancelScan();

        bool IsFeeder();

        // get preview image
        // 
        void GetPreview();

        // do scan
        HRESULT Scan(
            const std::wstring& saveDirectory, 
            const std::wstring& saveFilename, 
            std::vector<std::wstring>& saveFilenames,
            ScanProgressCallback progressCallback = nullptr);

    private:
        // Build WIA item tree from a IWiaItem pointer
        static std::shared_ptr<WIAItemTreeNode> FindImageSourcesFromDevice(ATL::CComPtr<IWiaItem2> pDevice);
        static std::shared_ptr<WIAItemTreeNode> DoFindImageSourcesFromDevice(ATL::CComPtr<IWiaItem2> pParentDevice);
        // Get the pointer of the image source with a specific name from the WIA item tree.
        // If the name is empty, Get The first image source pointer from the WIA item tree which can be used for image transfer.
        static ATL::CComPtr<IWiaItem2> GetImageSource(std::shared_ptr<WIAItemTreeNode> itemTree, const std::wstring& imgSourceName);

        bool SetDeviceDocumentHandling(ATL::CComPtr<IWiaItem2> device, const std::wstring& handling);
        bool SetDeviceImageFormat(ATL::CComPtr<IWiaItem2> device, const std::wstring& imageFormat);
        bool SetDeviceScanPageCount(ATL::CComPtr<IWiaItem2> device, int pageCount);

    private:
        std::recursive_mutex m_lockWIADevice;

        CWIADeviceMgr& m_manager;
        ATL::CComPtr<IWiaItem2> m_pDevice;
        DWORD m_deviceCookie;


        std::shared_ptr<WIAItemTreeNode> m_imageSources;

        bool m_bScanRunning;

        // Settings set when the scan operation is going to start
        std::wstring m_documentHandling; // document handling
        std::wstring m_imageFormat;   // output file format
        int m_pageCount;           // how many pages will be scanned
    };
}

