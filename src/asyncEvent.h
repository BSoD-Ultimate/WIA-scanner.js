#pragma once
#include <uv.h>

namespace scanner
{
    class uvAsyncEvent
    {
    public:
        uvAsyncEvent(void* context, uv_async_cb callback);
        virtual ~uvAsyncEvent();
        void* GetContext() const;
        void NotifyComplete();

    private:
        static void uvCloseCallback(uv_handle_t* handle)
        {
            delete (uv_async_t*)handle;
        }
    protected:
        struct uvAsyncEventDeleter
        {
            void operator()(uv_async_t* asyncEvent) const
            {
                if (asyncEvent)
                {
                    uv_close((uv_handle_t*)asyncEvent, uvCloseCallback);
                }
            }
        };
        std::unique_ptr<uv_async_t, uvAsyncEventDeleter> m_pAsyncEvent;
    };
}