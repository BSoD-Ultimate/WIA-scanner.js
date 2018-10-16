#include "stdafx.h"
#include "asyncEvent.h"

namespace scanner
{
    uvAsyncEvent::uvAsyncEvent(void * context, uv_async_cb callback)
    {
        m_pAsyncEvent.reset(new uv_async_t());
        m_pAsyncEvent->data = context;
        uv_async_init(uv_default_loop(), m_pAsyncEvent.get(), callback);
    }

    uvAsyncEvent::~uvAsyncEvent()
    {
    }

    void* uvAsyncEvent::GetContext() const
    {
        return m_pAsyncEvent->data;
    }

    void uvAsyncEvent::NotifyComplete()
    {
        uv_async_send(m_pAsyncEvent.get());
    }
}