#include "stdafx.h"

#include <nan.h>

#include "WIADeviceMgr.h"
#include "asyncEvent.h"

#include <experimental/filesystem>

#define CHECK_VALUE_TYPE(value, type, errMsg) \
    if(!value->Is##type()) \
    { \
        Nan::ThrowTypeError(errMsg); \
        return; \
    }

namespace scanner
{
    // global variables
    static bool g_bInit = false;
    // global COM environment
    static std::unique_ptr<util::COMEnvironment> g_comEnvironment;


    static void cleanup();


    // Wrap WIA device handle to JavaScript
    class WIADeviceJSWrap
        : public Nan::ObjectWrap
    {
    public:
        static NAN_MODULE_INIT(Init);

        std::shared_ptr<CWIADevice> GetDevice() const;
    private:
        explicit WIADeviceJSWrap(std::shared_ptr<CWIADevice> device);
        ~WIADeviceJSWrap();

        static NAN_METHOD(NewInstance);

        static NAN_METHOD(Close);

        static NAN_METHOD(ShowDeviceDlg);
        static NAN_METHOD(SetCallback);
        static NAN_METHOD(GetSources);
        static NAN_METHOD(GetProperties);
        static NAN_METHOD(SetProperties);
        static NAN_METHOD(GetPreview);
        static NAN_METHOD(IsFeeder);
        static NAN_METHOD(DoScan);
        static NAN_METHOD(Cancel);

    private:
        std::shared_ptr<CWIADevice> m_device;

    private:
        std::shared_ptr<Nan::Callback> m_pScanCompleteCallback;
        std::shared_ptr<Nan::Callback> m_pScanProgressCallback;
    private:

        static v8::Persistent<v8::Function> constructor;
    };
    v8::Persistent<v8::Function> WIADeviceJSWrap::constructor;

    NAN_MODULE_INIT(WIADeviceJSWrap::Init)
    {
        v8::Isolate* isolate = target->GetIsolate();
        // Prepare constructor template
        v8::Local<v8::FunctionTemplate> tpl = Nan::New<v8::FunctionTemplate>(NewInstance);
        tpl->SetClassName(v8::String::NewFromUtf8(isolate, "WIADevice"));
        tpl->InstanceTemplate()->SetInternalFieldCount(1);

        // Prototype

        Nan::SetPrototypeMethod(tpl, "close", Close);

        Nan::SetPrototypeMethod(tpl, "showDeviceDlg", ShowDeviceDlg);
        Nan::SetPrototypeMethod(tpl, "on", SetCallback);
        Nan::SetPrototypeMethod(tpl, "getSources", GetSources);
        Nan::SetPrototypeMethod(tpl, "getProperties", GetProperties);
        Nan::SetPrototypeMethod(tpl, "setProperties", SetProperties);
        Nan::SetPrototypeMethod(tpl, "getPreview", GetPreview);
        Nan::SetPrototypeMethod(tpl, "isFeeder", IsFeeder);
        Nan::SetPrototypeMethod(tpl, "doScan", DoScan);
        Nan::SetPrototypeMethod(tpl, "cancel", Cancel);

        constructor.Reset(isolate, tpl->GetFunction());
        target->Set(Nan::New("WIADevice").ToLocalChecked(), tpl->GetFunction());
    }



    std::shared_ptr<CWIADevice> WIADeviceJSWrap::GetDevice() const
    {
        return m_device;
    }

    WIADeviceJSWrap::WIADeviceJSWrap(std::shared_ptr<CWIADevice> device)
        : m_device(device)
    {
        assert(m_device);
    }
    WIADeviceJSWrap::~WIADeviceJSWrap()
    {

    }

    NAN_METHOD(WIADeviceJSWrap::NewInstance)
    {
        v8::Isolate* isolate = info.GetIsolate();

        if (info.IsConstructCall()) {
            // Invoked as constructor: `new MyObject(...)`
            CHECK_VALUE_TYPE(info[0], String, "type \"string\" expected in argument 1.");

            v8::Local<v8::String> deviceUUIDValue = v8::Local<v8::String>::Cast(info[0]);
            std::wstring deviceUUID = util::WStringFromUTF8(*v8::String::Utf8Value(deviceUUIDValue));

            try
            {
                auto device = CWIADeviceMgr::GetInstance()->OpenWIADevice(deviceUUID);

                if (!device)
                {
                    throw std::runtime_error("unable to open WIA device");
                }

                WIADeviceJSWrap* obj = new WIADeviceJSWrap(device);
                obj->Wrap(info.This());
                info.GetReturnValue().Set(info.This());
            }
            catch (const std::exception&)
            {
                char errBuf[500] = { 0 };
                sprintf_s(errBuf, "Unable to open device \"%s\".", util::WStringToUTF8(deviceUUID).c_str());
                Nan::ThrowError(errBuf);
                return;
            }
        }
        else {
            // Invoked as plain function `MyObject(...)`, turn into construct call.
            const int argc = 1;
            v8::Local<v8::Value> argv[argc] = { info[0] };
            v8::Local<v8::Context> context = isolate->GetCurrentContext();
            v8::Local<v8::Function> cons = v8::Local<v8::Function>::New(isolate, constructor);
            v8::Local<v8::Object> result =
                cons->NewInstance(context, argc, argv).ToLocalChecked();
            info.GetReturnValue().Set(result);
        }
    }

    NAN_METHOD(WIADeviceJSWrap::Close)
    {
        v8::Isolate* isolate = info.GetIsolate();
        WIADeviceJSWrap* obj = Nan::ObjectWrap::Unwrap<WIADeviceJSWrap>(info.Holder());
        assert(obj);

        if (!obj->GetDevice())
        {
            assert(false);
            Nan::ThrowError("Object has been destroyed");
            return;
        }

        obj->GetDevice().reset();
    }

    NAN_METHOD(WIADeviceJSWrap::SetCallback)
    {
        v8::Isolate* isolate = info.GetIsolate();
        WIADeviceJSWrap* obj = Nan::ObjectWrap::Unwrap<WIADeviceJSWrap>(info.Holder());

        assert(obj);

        CHECK_VALUE_TYPE(info[0], String, "type \"string\" expected in argument 1.");

        v8::Local<v8::String> callbackTypeValue = v8::Local<v8::String>::Cast(info[0]);
        std::string callbackType = *v8::String::Utf8Value(callbackTypeValue);

        std::shared_ptr<Nan::Callback> callbk;

        if (!info[1]->IsNullOrUndefined())
        {
            CHECK_VALUE_TYPE(info[1], Function, "type \"function\" expected in argument 2.");
            callbk = std::shared_ptr<Nan::Callback>(new Nan::Callback(Nan::To<v8::Function>(info[1]).ToLocalChecked()));
        }

        if (!obj->GetDevice())
        {
            assert(false);
            Nan::ThrowError("Object has been destroyed");
            return;
        }

        if (callbackType == "progress")
        {
            obj->m_pScanProgressCallback = callbk;
        }
        else if (callbackType == "complete")
        {
            obj->m_pScanCompleteCallback = callbk;
        }

    }

    NAN_METHOD(WIADeviceJSWrap::ShowDeviceDlg)
    {
        v8::Isolate* isolate = info.GetIsolate();
        WIADeviceJSWrap* obj = Nan::ObjectWrap::Unwrap<WIADeviceJSWrap>(info.Holder());
        assert(obj);

        HWND parentHWND = 0;

        // HWND handle
        if (!info[0]->IsNullOrUndefined())
        {
            CHECK_VALUE_TYPE(info[0], Object, "type \"Object\" expected in argument 1.");
            parentHWND = *(HWND*)node::Buffer::Data(info[1]);
        }
        CHECK_VALUE_TYPE(info[1], String, "type \"string\" expected in argument 2.");
        CHECK_VALUE_TYPE(info[2], String, "type \"string\" expected in argument 3.");

        v8::Local<v8::String> recvFolderValue = v8::Local<v8::String>::Cast(info[1]);
        v8::Local<v8::String> filenameTemplateValue = v8::Local<v8::String>::Cast(info[2]);

        std::wstring recvFolder = util::WStringFromUTF8(*v8::String::Utf8Value(recvFolderValue));
        std::wstring filenameTemplate = util::WStringFromUTF8(*v8::String::Utf8Value(filenameTemplateValue));

        if (!obj->GetDevice())
        {
            assert(false);
            Nan::ThrowError("Object has been destroyed");
            return;
        }

        // Open the scan dialog provided by Windows. Do modal loop here
        std::vector<std::wstring> outFilepaths;
        HRESULT hr = obj->GetDevice()->ShowDeviceDlg(parentHWND, recvFolder, filenameTemplate, outFilepaths);

        v8::Local<v8::Object> retObject = Nan::New<v8::Object>();
        v8::Local<v8::Array> filesArray = Nan::New<v8::Array>();

        for (size_t i = 0; i < outFilepaths.size(); i++)
        {
            filesArray->Set(i, Nan::New(util::WStringToUTF8(outFilepaths[i])).ToLocalChecked());
        }

        retObject->Set(Nan::New("files").ToLocalChecked(), filesArray);
        retObject->Set(Nan::New("retCode").ToLocalChecked(), Nan::New(hr));
        retObject->Set(Nan::New("errorMsg").ToLocalChecked(), Nan::New(util::WStringToUTF8(util::GetWIAErrorStr(hr))).ToLocalChecked());

        info.GetReturnValue().Set(retObject);
    }

    NAN_METHOD(WIADeviceJSWrap::GetSources)
    {
        v8::Isolate* isolate = info.GetIsolate();
        WIADeviceJSWrap* obj = Nan::ObjectWrap::Unwrap<WIADeviceJSWrap>(info.Holder());
        assert(obj);

        if (!obj->GetDevice())
        {
            assert(false);
            Nan::ThrowError("Object has been destroyed");
            return;
        }

        auto sources = obj->GetDevice()->GetImageSources();

        v8::Local<v8::Object> retObject = Nan::New<v8::Object>();
        v8::Local<v8::Array> sourcesArray = Nan::New<v8::Array>();

        for (size_t i = 0; i < sources.size(); i++)
        {
            sourcesArray->Set(i, Nan::New(util::WStringToUTF8(sources[i])).ToLocalChecked());
        }

        retObject->Set(Nan::New("sources").ToLocalChecked(), sourcesArray);

        info.GetReturnValue().Set(retObject);
    }

    NAN_METHOD(WIADeviceJSWrap::GetProperties)
    {
        v8::Isolate* isolate = info.GetIsolate();
        WIADeviceJSWrap* obj = Nan::ObjectWrap::Unwrap<WIADeviceJSWrap>(info.Holder());

        assert(obj);

        if (!obj->GetDevice())
        {
            assert(false);
            Nan::ThrowError("Object has been destroyed");
            return;
        }

        v8::Local<v8::Object> retObject = Nan::New<v8::Object>();
        
        // output file format
        std::wstring imageFormat = obj->GetDevice()->GetImageFormat();
        retObject->Set(Nan::New("format").ToLocalChecked(), Nan::New(util::WStringToUTF8(imageFormat)).ToLocalChecked());

        // paper profile
        std::wstring paperProfile = obj->GetDevice()->GetPaperProfile();
        retObject->Set(Nan::New("paper").ToLocalChecked(), Nan::New(util::WStringToUTF8(paperProfile)).ToLocalChecked());

        // color mode
        std::wstring color = obj->GetDevice()->GetColorFormat();
        retObject->Set(Nan::New("color").ToLocalChecked(), Nan::New(util::WStringToUTF8(color)).ToLocalChecked());

        // brightness range
        int min, max, normal, step;
        obj->GetDevice()->GetScanBrightnessRange(min, max, normal, step);
        v8::Local<v8::Object> brightnessRangeObject = Nan::New<v8::Object>();
        brightnessRangeObject->Set(Nan::New("min").ToLocalChecked(), Nan::New(min/step));
        brightnessRangeObject->Set(Nan::New("max").ToLocalChecked(), Nan::New(max/step));
        retObject->Set(Nan::New("brightness_range").ToLocalChecked(), brightnessRangeObject);

        // contrast range
        obj->GetDevice()->GetScanContrastRange(min, max, normal, step);
        v8::Local<v8::Object> contrastRangeObject = Nan::New<v8::Object>();
        contrastRangeObject->Set(Nan::New("min").ToLocalChecked(), Nan::New(min / step));
        contrastRangeObject->Set(Nan::New("max").ToLocalChecked(), Nan::New(max / step));
        retObject->Set(Nan::New("contrast_range").ToLocalChecked(), contrastRangeObject);

        // brightness
        int brightness = obj->GetDevice()->GetScanBrightness();
        retObject->Set(Nan::New("brightness").ToLocalChecked(), Nan::New(brightness));

        // contrast
        int contrast = obj->GetDevice()->GetScanContrast();
        retObject->Set(Nan::New("contrast").ToLocalChecked(), Nan::New(contrast));

        // DPI
        int dpi = obj->GetDevice()->GetScanDPI();
        retObject->Set(Nan::New("dpi").ToLocalChecked(), Nan::New(dpi));

        // page count
        int pageCount = obj->GetDevice()->GetScanPageCount();
        if (pageCount)
        {
            retObject->Set(Nan::New("pageCount").ToLocalChecked(), Nan::New(pageCount));
        }
        else
        {
            retObject->Set(Nan::New("pageCount").ToLocalChecked(), Nan::New("all").ToLocalChecked());
        }

        // document handling
        std::wstring handling = obj->GetDevice()->GetDocumentHandling();
        retObject->Set(Nan::New("document_handling").ToLocalChecked(), Nan::New(util::WStringToUTF8(handling)).ToLocalChecked());

        info.GetReturnValue().Set(retObject);
    }
    NAN_METHOD(WIADeviceJSWrap::SetProperties)
    {
        v8::Isolate* isolate = info.GetIsolate();
        WIADeviceJSWrap* obj = Nan::ObjectWrap::Unwrap<WIADeviceJSWrap>(info.Holder());
        assert(obj);

        if (!obj->GetDevice())
        {
            assert(false);
            Nan::ThrowError("Object has been destroyed");
            return;
        }

        CHECK_VALUE_TYPE(info[0], Object, "type \"object\" expected in argument 1.");

        v8::Local<v8::Object> retObject = Nan::New<v8::Object>();

        v8::Local<v8::Object> paramObj = v8::Local<v8::Object>::Cast(info[0]);

        // output file format
        {
            v8::Local<v8::String> imageFormatValue = v8::Local<v8::String>::Cast(paramObj->Get(Nan::New("format").ToLocalChecked()));
            if (!imageFormatValue->IsNullOrUndefined())
            {
                CHECK_VALUE_TYPE(imageFormatValue, String, "type \"string\" expected in value \"format\"");
                std::wstring imgFormat = util::WStringFromUTF8(*v8::String::Utf8Value(imageFormatValue));
                bool ret = obj->GetDevice()->SetImageFormat(imgFormat);

                retObject->Set(Nan::New("format").ToLocalChecked(), Nan::New(ret));
            }
        }

        // paper profile
        {
            v8::Local<v8::String> paperSizeValue = v8::Local<v8::String>::Cast(paramObj->Get(Nan::New("paper").ToLocalChecked()));
            if (!paperSizeValue->IsNullOrUndefined())
            {
                CHECK_VALUE_TYPE(paperSizeValue, String, "type \"string\" expected in value \"paper\"");

                std::wstring paperProfile = util::WStringFromUTF8(*v8::String::Utf8Value(paperSizeValue));
                obj->GetDevice()->GetPaperProfile();
                bool ret = obj->GetDevice()->SetPaperProfile(paperProfile);

                retObject->Set(Nan::New("paper").ToLocalChecked(), Nan::New(ret));
            }
        }

        // color mode
        {
            v8::Local<v8::String> colorValue = v8::Local<v8::String>::Cast(paramObj->Get(Nan::New("color").ToLocalChecked()));
            if (!colorValue->IsNullOrUndefined())
            {
                CHECK_VALUE_TYPE(colorValue, String, "type \"string\" expected in value \"color\"");
                std::wstring imgFormat = util::WStringFromUTF8(*v8::String::Utf8Value(colorValue));
                bool ret = obj->GetDevice()->SetColorFormat(imgFormat);

                retObject->Set(Nan::New("color").ToLocalChecked(), Nan::New(ret));
            }
        }

        // brightness
        {
            v8::Local<v8::Integer> brightnessValue = v8::Local<v8::Integer>::Cast(paramObj->Get(Nan::New("brightness").ToLocalChecked()));
            if (!brightnessValue->IsNullOrUndefined())
            {
                CHECK_VALUE_TYPE(brightnessValue, Number, "type \"number\" expected in value \"brightness\"");
                int brightness = brightnessValue->IntegerValue();

                bool ret = obj->GetDevice()->SetScanBrightness(brightness);

                retObject->Set(Nan::New("brightness").ToLocalChecked(), Nan::New(ret));
            }
        }

        // contrast
        {
            v8::Local<v8::Integer> contrastValue = v8::Local<v8::Integer>::Cast(paramObj->Get(Nan::New("contrast").ToLocalChecked()));
            if (!contrastValue->IsNullOrUndefined())
            {
                CHECK_VALUE_TYPE(contrastValue, Number, "type \"number\" expected in value \"contrast\"");
                int contrast = contrastValue->IntegerValue();
                bool ret = obj->GetDevice()->SetScanContrast(contrast);

                retObject->Set(Nan::New("contrast").ToLocalChecked(), Nan::New(ret));
            }
        }

        // DPI
        {
            v8::Local<v8::Integer> dpiValue = v8::Local<v8::Integer>::Cast(paramObj->Get(Nan::New("dpi").ToLocalChecked()));
            if (!dpiValue->IsNullOrUndefined())
            {
                CHECK_VALUE_TYPE(dpiValue, Number, "type \"number\" expected in value \"dpi\"");
                int dpi = dpiValue->IntegerValue();
                bool ret = obj->GetDevice()->SetScanDPI(dpi);

                retObject->Set(Nan::New("dpi").ToLocalChecked(), Nan::New(ret));
            }
        }

        // page count
        {
            v8::Local<v8::Value> pageCountValue = paramObj->Get(Nan::New("pageCount").ToLocalChecked());
            if (!pageCountValue->IsNullOrUndefined())
            {
                int pageCount = 0;
                if (pageCountValue->IsString())
                {
                    std::string count = *v8::String::Utf8Value(v8::Local<v8::String>::Cast(pageCountValue));
                    if (count == "all")
                    {
                        pageCount = 0;
                    }
                }
                else if (pageCountValue->IsNumber())
                {
                    pageCount = pageCountValue->IntegerValue();
                }
                else
                {
                    Nan::ThrowTypeError("type \"number\" or \"string\" expected in value \"pageCount\"");
                    return;
                }

                bool ret = obj->GetDevice()->SetScanPageCount(pageCount);

                retObject->Set(Nan::New("pageCount").ToLocalChecked(), Nan::New(ret));
            }
        }

        // document handling
        {
            v8::Local<v8::String> docHandlingValue = v8::Local<v8::String>::Cast(paramObj->Get(Nan::New("document_handling").ToLocalChecked()));
            if (!docHandlingValue->IsNullOrUndefined())
            {
                CHECK_VALUE_TYPE(docHandlingValue, String, "type \"string\" expected in value \"format\"");
                std::wstring handling = util::WStringFromUTF8(*v8::String::Utf8Value(docHandlingValue));
                bool ret = obj->GetDevice()->SetDocumentHandling(handling);

                retObject->Set(Nan::New("document_handling").ToLocalChecked(), Nan::New(ret));
            }
        }

        info.GetReturnValue().Set(retObject);
    }

    NAN_METHOD(WIADeviceJSWrap::GetPreview)
    {
        v8::Isolate* isolate = info.GetIsolate();
        WIADeviceJSWrap* obj = Nan::ObjectWrap::Unwrap<WIADeviceJSWrap>(info.Holder());
        assert(obj);

        if (!obj->GetDevice())
        {
            assert(false);
            Nan::ThrowError("Object has been destroyed");
            return;
        }

    }

    NAN_METHOD(WIADeviceJSWrap::IsFeeder)
    {
        v8::Isolate* isolate = info.GetIsolate();
        WIADeviceJSWrap* obj = Nan::ObjectWrap::Unwrap<WIADeviceJSWrap>(info.Holder());
        assert(obj);

        if (!obj->GetDevice())
        {
            assert(false);
            Nan::ThrowError("Object has been destroyed");
            return;
        }

        info.GetReturnValue().Set(Nan::New(obj->GetDevice()->IsFeeder()));
    }

    NAN_METHOD(WIADeviceJSWrap::DoScan)
    {
        v8::Isolate* isolate = info.GetIsolate();
        WIADeviceJSWrap* obj = Nan::ObjectWrap::Unwrap<WIADeviceJSWrap>(info.Holder());
        assert(obj);

        if (!obj->GetDevice())
        {
            assert(false);
            Nan::ThrowError("Object has been destroyed");
            return;
        }

        CHECK_VALUE_TYPE(info[0], Object, "type \"object\" expected in argument 1.");

        if (!info[1]->IsNullOrUndefined())
        {
            CHECK_VALUE_TYPE(info[1], Function, "type \"function\" expected in argument 2.");
            std::shared_ptr<Nan::Callback> callbk(new Nan::Callback(Nan::To<v8::Function>(info[1]).ToLocalChecked()));
            obj->m_pScanCompleteCallback = callbk;
        }

        v8::Local<v8::Object> paramObj = v8::Local<v8::Object>::Cast(info[0]);

        v8::Local<v8::String> saveDirValue = v8::Local<v8::String>::Cast(paramObj->Get(Nan::New("saveDir").ToLocalChecked()));
        v8::Local<v8::String> saveFilenameValue = v8::Local<v8::String>::Cast(paramObj->Get(Nan::New("saveFilename").ToLocalChecked()));

        CHECK_VALUE_TYPE(saveDirValue, String, "type \"string\" expected in value \"saveDir\".");
        CHECK_VALUE_TYPE(saveFilenameValue, String, "type \"string\" expected in value \"saveFilename\".");

        std::wstring saveDir = util::WStringFromUTF8(*v8::String::Utf8Value(saveDirValue));
        std::wstring saveFilename = util::WStringFromUTF8(*v8::String::Utf8Value(saveFilenameValue));

        class ScanWorker : public Nan::AsyncWorker
        {
        public:
            ScanWorker(WIADeviceJSWrap* obj, const std::wstring& saveDir, const std::wstring& saveFilename)
                : Nan::AsyncWorker(NULL)
                , m_pObj(obj)
                , m_saveDir(saveDir)
                , m_saveFilename(saveFilename)
                , m_pProgressEvent(new uvAsyncEvent(this, progressCallback))
                , m_hrScanResult(S_OK)
            {
            }

            void Execute() override
            {
                util::COMEnvironment env;
                // firstly, create directory if needed
                std::experimental::filesystem::create_directories(m_saveDir);

                m_hrScanResult = m_pObj->GetDevice()->Scan(m_saveDir, m_saveFilename, m_saveFilePaths,
                    [this](const ScanProgressInfo& info)
                {
                    std::lock_guard<std::recursive_mutex> g(m_lockProgress);
                    m_progressInfo = info;
                    
                    m_pProgressEvent->NotifyComplete();
                });
            }

            void HandleOKCallback() override
            {
                Nan::HandleScope scope;

                if (m_pObj->m_pScanCompleteCallback)
                {
                    Nan::HandleScope scope;

                    v8::Local<v8::Object> retObject = Nan::New<v8::Object>();

                    retObject->Set(Nan::New("retCode").ToLocalChecked(), Nan::New(m_hrScanResult));
                    std::wstring errorMsg = util::GetWIAErrorStr(m_hrScanResult);
                    retObject->Set(Nan::New("errMsg").ToLocalChecked(), Nan::New(util::WStringToUTF8(errorMsg)).ToLocalChecked());

                    v8::Local<v8::Array> filesArray = Nan::New<v8::Array>();
                    for (int i = 0; i < m_saveFilePaths.size(); i++)
                    {
                        filesArray->Set(i, Nan::New(util::WStringToUTF8(m_saveFilePaths[i])).ToLocalChecked());
                    }
                    retObject->Set(Nan::New("files").ToLocalChecked(), filesArray);

                    int argc = 1;
                    std::unique_ptr<v8::Local<v8::Value>[]> argv(new v8::Local<v8::Value>[argc]());
                    argv[0] = retObject;
                    Nan::Call(*m_pObj->m_pScanCompleteCallback, argc, argv.get());
                }
            }

        private:
            static void progressCallback(uv_async_t* handle)
            {
                auto* pThis = reinterpret_cast<ScanWorker*>(handle->data);

                std::lock_guard<std::recursive_mutex> g(pThis->m_lockProgress);
                if (pThis->m_pObj->m_pScanProgressCallback)
                {
                    Nan::HandleScope scope;

                    v8::Local<v8::Object> retObject = Nan::New<v8::Object>();

                    retObject->Set(Nan::New("page").ToLocalChecked(), Nan::New(pThis->m_progressInfo.pageCount));
                    retObject->Set(Nan::New("percent").ToLocalChecked(), Nan::New(pThis->m_progressInfo.percentComplete));
                    retObject->Set(Nan::New("bytesTransferred").ToLocalChecked(), Nan::New(double(pThis->m_progressInfo.bytesTransferred)));
                    retObject->Set(Nan::New("error").ToLocalChecked(), Nan::New(pThis->m_progressInfo.error));
                    retObject->Set(Nan::New("errMsg").ToLocalChecked(), Nan::New(util::WStringToUTF8(util::GetWIAErrorStr(pThis->m_progressInfo.error))).ToLocalChecked());

                    int argc = 1;
                    std::unique_ptr<v8::Local<v8::Value>[]> argv(new v8::Local<v8::Value>[argc]());
                    argv[0] = retObject;
                    Nan::Call(*pThis->m_pObj->m_pScanProgressCallback, argc, argv.get());
                }
            }

        private:
            WIADeviceJSWrap* m_pObj;
            std::wstring m_saveDir;
            std::wstring m_saveFilename;

            // members for progress info
            std::unique_ptr<uvAsyncEvent> m_pProgressEvent;
            std::recursive_mutex m_lockProgress;
            ScanProgressInfo m_progressInfo;

            HRESULT m_hrScanResult;
            // paths of the acquired image files
            std::vector<std::wstring> m_saveFilePaths;
        };
        ScanWorker* worker = new ScanWorker(obj, saveDir, saveFilename);
        Nan::AsyncQueueWorker(worker);
    }

    NAN_METHOD(WIADeviceJSWrap::Cancel)
    {
        v8::Isolate* isolate = info.GetIsolate();
        WIADeviceJSWrap* obj = Nan::ObjectWrap::Unwrap<WIADeviceJSWrap>(info.Holder());

        assert(obj);

        if (!obj->GetDevice())
        {
            assert(false);
            Nan::ThrowError("Object has been destroyed");
            return;
        }

        obj->GetDevice()->CancelScan();

    }


    static NAN_METHOD(ListAllDevices)
    {
        auto devices = CWIADeviceMgr::GetInstance()->ListAllDevices();

        v8::Local<v8::Array> retDevicesInfo = Nan::New<v8::Array>();

        for (size_t i = 0; i < devices.size(); i++)
        {
            v8::Local<v8::Object> deviceInfo = Nan::New<v8::Object>();

            deviceInfo->Set(Nan::New("deviceUUID").ToLocalChecked(), Nan::New(util::WStringToUTF8(devices[i]->deviceUUID)).ToLocalChecked());
            deviceInfo->Set(Nan::New("manufacturer").ToLocalChecked(), Nan::New(util::WStringToUTF8(devices[i]->manufacturer)).ToLocalChecked());
            deviceInfo->Set(Nan::New("description").ToLocalChecked(), Nan::New(util::WStringToUTF8(devices[i]->description)).ToLocalChecked());
            deviceInfo->Set(Nan::New("deviceType").ToLocalChecked(), Nan::New(devices[i]->type));
            deviceInfo->Set(Nan::New("portName").ToLocalChecked(), Nan::New(util::WStringToUTF8(devices[i]->port)).ToLocalChecked());
            deviceInfo->Set(Nan::New("deviceName").ToLocalChecked(), Nan::New(util::WStringToUTF8(devices[i]->deviceName)).ToLocalChecked());
            deviceInfo->Set(Nan::New("serverName").ToLocalChecked(), Nan::New(util::WStringToUTF8(devices[i]->server)).ToLocalChecked());
            deviceInfo->Set(Nan::New("remoteDeviceId").ToLocalChecked(), Nan::New(util::WStringToUTF8(devices[i]->remoteDeviceId)).ToLocalChecked());
            deviceInfo->Set(Nan::New("uiCLSID").ToLocalChecked(), Nan::New(util::WStringToUTF8(devices[i]->uiClassId)).ToLocalChecked());
            deviceInfo->Set(Nan::New("hardwareConfig").ToLocalChecked(), Nan::New(devices[i]->hardwareConfig));
            deviceInfo->Set(Nan::New("baudrate").ToLocalChecked(), Nan::New(util::WStringToUTF8(devices[i]->baudrate)).ToLocalChecked());
            deviceInfo->Set(Nan::New("STIGenericCapabilities").ToLocalChecked(), Nan::New(devices[i]->STIGenericCapabilities));
            deviceInfo->Set(Nan::New("wiaVersion").ToLocalChecked(), Nan::New(util::WStringToUTF8(devices[i]->WIAVersion)).ToLocalChecked());
            deviceInfo->Set(Nan::New("driverVersion").ToLocalChecked(), Nan::New(util::WStringToUTF8(devices[i]->DriverVersion)).ToLocalChecked());
            deviceInfo->Set(Nan::New("PnPIDString").ToLocalChecked(), Nan::New(util::WStringToUTF8(devices[i]->PnPIDString)).ToLocalChecked());
            deviceInfo->Set(Nan::New("STIDriverVersion").ToLocalChecked(), Nan::New(devices[i]->STIDriverVersion));

            retDevicesInfo->Set(i, deviceInfo);
        }

        info.GetReturnValue().Set(retDevicesInfo);
    }

    static NAN_METHOD(OpenDevice)
    {
        v8::Isolate* isolate = info.GetIsolate();
        CHECK_VALUE_TYPE(info[0], String, "type \"string\" excepted in argument 1.");

        v8::Local<v8::String> uuidValue = v8::Local<v8::String>::Cast(info[0]);

        std::wstring deviceUUID = util::WStringFromUTF8(*v8::String::Utf8Value(uuidValue));
    }

    static NAN_METHOD(Cleanup)
    {
        v8::Isolate* isolate = info.GetIsolate();
        cleanup();
    }

    static void InitCOMEnvironment()
    {
        // We may already have a COM environment available here.
        // For Example, Electron Main process initializes a COM environment before the library loads.
        // The environment initialized by Electron is APTTYPE_MAINSTA(Main, single-threaded).
        // We cannot determine The COM threading model here.
        // if the COM environment is not available, the threading model will be initialized as multi-threaded.
        APTTYPE apartmentType;
        APTTYPEQUALIFIER aptTypeQualifier;
        HRESULT hr = CoGetApartmentType(&apartmentType, &aptTypeQualifier);

        if (FAILED(hr))
        {
            g_comEnvironment.reset(new util::COMEnvironment());
        }
    }

    static void ShutdownComEnvironment()
    {
        // Shutdown the COM environment initialized by the library
        if (g_comEnvironment)
        {
            g_comEnvironment.reset();
        }
    }

    static void cleanup()
    {
        CWIADeviceMgr::GetInstance().reset();

        ShutdownComEnvironment();
        g_bInit = false;
    }

    static void atExit(void* arg)
    {
        if (g_bInit)
        {
            cleanup();
        }
    }
}

NAN_MODULE_INIT(init)
{
    using namespace scanner;

    node::AtExit(atExit, NULL);
    // Initialize COM environment if needed
    InitCOMEnvironment();

    // Initialize the global object IWiaDevMgr, throws error if fails.
    try
    {
        auto& pDeviceMgr = scanner::CWIADeviceMgr::GetInstance();
    }
    catch (const std::exception&)
    {
        Nan::ThrowError("Unable to initialize WIA device manager! module init failed.");
        return;
    }

    scanner::WIADeviceJSWrap::Init(target);

    Nan::SetMethod(target, "listAllDevices", ListAllDevices);
    Nan::SetMethod(target, "openDevice", OpenDevice);
    Nan::SetMethod(target, "cleanup", Cleanup);

    g_bInit = true;
}

NODE_MODULE(scanner_nodejs, init)