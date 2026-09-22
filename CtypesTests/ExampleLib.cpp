/** Shared library with the example COM classes, so that a client outside C++ has
    something to activate. Only the CLSID and the IDL description are needed to
    reach it, the same as for any other COM server. */

#include <AppAPI/ComSupport.hpp>
#include <algorithm>
#include <cwctype>
#include <sstream>
#include <string>
#include "Registry.h"

#ifndef VARIANT_TRUE
#define VARIANT_TRUE  ((VARIANT_BOOL)-1)
#define VARIANT_FALSE ((VARIANT_BOOL)0)
#endif


/** Implementation of the Calculator coclass. */
class CalculatorImpl :
    public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public CComCoClass<CalculatorImpl>,
    public ICalcExt, public ICalc2 {
public:
    HRESULT GetValue (/*out*/int * value) override {
        if (!value)
            return E_INVALIDARG;

        *value = 42;
        return S_OK;
    }

    HRESULT Add (int a, int b, /*out*/int * result) override {
        if (!result)
            return E_INVALIDARG;

        *result = a + b;
        return S_OK;
    }

    HRESULT SetCallback (ICalcCb * cb) override {
        if (!cb)
            return E_INVALIDARG;

        m_callback = cb;
        return S_OK;
    }

    HRESULT GetValue2 (/*out*/int * value) override {
        if (!value)
            return E_INVALIDARG;

        *value = 43;

        if (m_callback)
            m_callback->Message(CComBSTR(L"GetValue2 called"));

        return S_OK;
    }

    BEGIN_COM_MAP(CalculatorImpl)
        COM_INTERFACE_ENTRY(ICalc)
        COM_INTERFACE_ENTRY(ICalcExt)
        COM_INTERFACE_ENTRY(ICalc2)
    END_COM_MAP()

private:
    CComPtr<ICalcCb> m_callback;
};
OBJECT_ENTRY_AUTO(CLSID_Calculator, CalculatorImpl)


/** Implementation of the Registry coclass. */
class RegistryImpl :
    public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public CComCoClass<RegistryImpl>,
    public IRegistry, public IStats {
public:
    HRESULT NewCalculator (/*out*/ICalcExt ** calc) override {
        return CoCreateInstance(CLSID_Calculator, nullptr, CLSCTX_ALL, IID_ICalcExt, (void**)calc);
    }

    HRESULT Store (IUnknown * obj) override {
        m_stored = obj;
        m_store_count++;
        return S_OK;
    }

    HRESULT Fetch (/*out*/IUnknown ** obj) override {
        if (!obj)
            return E_INVALIDARG;

        *obj = m_stored;
        if (*obj)
            (*obj)->AddRef();
        return S_OK;
    }

    HRESULT Describe (ISource * source, /*out*/BSTR * text) override {
        if (!source || !text)
            return E_INVALIDARG;

        HRESULT hr = source->Message(CComBSTR(L"describing"));
        if (FAILED(hr))
            return hr;

        // a method the bindings cannot pass answers E_NOTIMPL from Python
        if (source->Opaque(nullptr) != E_NOTIMPL)
            return E_UNEXPECTED;

        CComBSTR name;
        hr = source->Name(&name);
        if (FAILED(hr))
            return hr;

        CComPtr<ICalc2> calc;
        hr = source->Calculator(&calc);
        if (FAILED(hr))
            return hr;
        int value = 0;
        hr = calc->GetValue2(&value);
        if (FAILED(hr))
            return hr;

        Span span = {};
        hr = source->GetSpan(&span);
        if (FAILED(hr))
            return hr;

        VARIANT_BOOL accepted = VARIANT_FALSE;
        hr = source->Accept(calc, &accepted);
        if (FAILED(hr))
            return hr;
        if (accepted != VARIANT_TRUE)
            return E_UNEXPECTED;

        int low = 0, high = 0;
        hr = source->Range(&low, &high);
        if (FAILED(hr))
            return hr;

        double adjusted = 1.5;
        hr = source->Adjust(&adjusted);
        if (FAILED(hr))
            return hr;

        const float values[3] = {1, 2, 5};
        float normalized[3] = {};
        hr = source->Normalize(values, normalized);
        if (FAILED(hr))
            return hr;

        CComSafeArray<float> samples;
        hr = source->Samples(&samples.m_ptr);
        if (FAILED(hr))
            return hr;
        float samples_sum = 0;
        for (unsigned int i = 0; i < samples.GetCount(); i++)
            samples_sum += samples.GetAt(i);

        CComSafeArray<BSTR> words;
        words.Add(CComBSTR(L"two"));
        words.Add(CComBSTR(L"words"));
        Record record = {};
        hr = source->Label(words, &record);
        if (FAILED(hr))
            return hr;
        CComBSTR label; // the record's string and array are ours to free
        label.Attach(record.name);
        CComSafeArray<double> lengths;
        if (record.samples)
            lengths.Attach(record.samples);
        if (record.shape != SHAPE_LINE || record.valid != VARIANT_TRUE || !record.flag || record.big != (1L << 40) || record.after != 2.5)
            return E_UNEXPECTED;

        std::wostringstream out;
        out << (const wchar_t*)name << L" uses " << value << L", range " << low << L"-" << high
            << L", adjusted " << adjusted << L", normalized " << normalized[0] << L" " << normalized[1] << L" " << normalized[2]
            << L", span " << span.low << L"-" << span.high << L", samples " << samples_sum
            << L", label " << (const wchar_t*)label << L" " << record.span.high << L" " << lengths.GetAt(0) << L"+" << lengths.GetAt(1);
        *text = CComBSTR(out.str().c_str()).Detach();
        return S_OK;
    }

    HRESULT MakeRecord (BSTR name, /*out*/Record * record) override {
        if (!record)
            return E_INVALIDARG;

        record->name = CComBSTR(name).Detach();
        record->shape = SHAPE_BOX;
        record->span.low = 1;
        record->span.high = 2;
        record->weights[0] = 0.5f;
        record->weights[1] = 0.25f;
        record->weights[2] = 0.25f;
        CComSafeArray<double> samples;
        samples.Add(1.5);
        samples.Add(2.5);
        record->samples = samples.Detach();
        record->valid = VARIANT_TRUE;
        record->flag = 1;
        record->big = 1L << 40;
        record->after = 2.5;
        return S_OK;
    }

    HRESULT CheckRecord (Record * record, /*out*/BSTR * summary) override {
        if (!record || !summary)
            return E_INVALIDARG;

        double sum = 0;
        unsigned int count = 0;
        if (record->samples) {
            CComSafeArray<double> samples; // borrowed, since the record stays the caller's
            samples.Attach(record->samples);
            count = samples.GetCount();
            for (unsigned int i = 0; i < count; i++)
                sum += samples.GetAt(i);
            samples.Detach();
        }

        std::wostringstream out;
        out << (record->name ? record->name : L"(null)") << L" " << (int)record->shape
            << L" " << record->span.low << L"-" << record->span.high
            << L" " << record->weights[0] << L"," << record->weights[1] << L"," << record->weights[2]
            << L" " << count << L":" << sum << L" " << (record->valid == VARIANT_TRUE ? L"valid" : L"invalid")
            << L" " << record->flag << L" " << record->big << L" " << record->after;
        *summary = CComBSTR(out.str().c_str()).Detach();
        return S_OK;
    }

    HRESULT Reverse (SAFEARRAY * data, /*out*/SAFEARRAY ** reversed) override {
        if (!data || !reversed)
            return E_INVALIDARG;

        CComSafeArray<BYTE> in;
        in.Attach(data);
        unsigned int count = in.GetCount();
        CComSafeArray<BYTE> out(count);
        for (unsigned int i = 0; i < count; i++)
            out.SetAt(i, in.GetAt(count-1-i));
        in.Detach(); // the caller's
        *reversed = out.Detach();
        return S_OK;
    }

    HRESULT Upper (SAFEARRAY * words, /*out*/SAFEARRAY ** upper) override {
        if (!words || !upper)
            return E_INVALIDARG;

        CComSafeArray<BSTR> in;
        in.Attach(words);
        CComSafeArray<BSTR> out;
        for (unsigned int i = 0; i < in.GetCount(); i++) {
            std::wstring word = (const wchar_t*)in.GetAt(i);
            for (wchar_t & c : word)
                c = towupper(c);
            out.Add(CComBSTR(word.c_str()));
        }
        in.Detach(); // the caller's
        *upper = out.Detach();
        return S_OK;
    }

    HRESULT Objects (/*out*/SAFEARRAY ** objects) override {
        if (!objects)
            return E_INVALIDARG;

        CComPtr<IUnknown> calc;
        HRESULT hr = CoCreateInstance(CLSID_Calculator, nullptr, CLSCTX_ALL, IID_IUnknown, (void**)&calc);
        if (FAILED(hr))
            return hr;

        CComSafeArray<IUnknown*> out;
        out.Add(m_stored);
        out.Add(calc);
        *objects = out.Detach();
        return S_OK;
    }

    HRESULT Opaque (void * /*data*/) override {
        return S_OK;
    }

    HRESULT GetSpan (const float values[3], /*out*/Span * span) override {
        if (!values || !span)
            return E_INVALIDARG;

        span->low = std::min(values[0], std::min(values[1], values[2]));
        span->high = std::max(values[0], std::max(values[1], values[2]));
        return S_OK;
    }

    HRESULT Sum (const float values[3], /*out*/double * sum) override {
        if (!values || !sum)
            return E_INVALIDARG;

        *sum = (double)values[0] + values[1] + values[2];
        return S_OK;
    }

    HRESULT StoreCount (/*out*/unsigned int * count) override {
        if (!count)
            return E_INVALIDARG;

        *count = m_store_count;
        return S_OK;
    }

    HRESULT MinMax (const float values[3], /*out*/float * min, /*out*/float * max) override {
        if (!values || !min || !max)
            return E_INVALIDARG;

        *min = std::min(values[0], std::min(values[1], values[2]));
        *max = std::max(values[0], std::max(values[1], values[2]));
        return S_OK;
    }

    HRESULT Bytes (unsigned int value, /*out*/unsigned char bytes[4]) override {
        if (!bytes)
            return E_INVALIDARG;

        for (int i = 0; i < 4; i++)
            bytes[i] = (value >> (8*i)) & 0xFF;
        return S_OK;
    }

    HRESULT Scale (/*in,out*/double * value, VARIANT_BOOL negate) override {
        if (!value)
            return E_INVALIDARG;

        *value *= 2;
        if (negate)
            *value = -*value;
        return S_OK;
    }

    HRESULT Positive (double value, /*out*/VARIANT_BOOL * positive) override {
        if (!positive)
            return E_INVALIDARG;

        *positive = (value > 0) ? VARIANT_TRUE : VARIANT_FALSE;
        return S_OK;
    }

    HRESULT Widen (Span span, float by, /*out*/Span * wider) override {
        if (!wider)
            return E_INVALIDARG;

        wider->low = span.low - by;
        wider->high = span.high + by;
        return S_OK;
    }

    HRESULT Next (Shape shape, /*out*/Shape * next) override {
        if (!next)
            return E_INVALIDARG;

        *next = (Shape)(shape << 1);
        return S_OK;
    }

    BEGIN_COM_MAP(RegistryImpl)
        COM_INTERFACE_ENTRY(IRegistry)
        COM_INTERFACE_ENTRY(IStats)
    END_COM_MAP()

private:
    CComPtr<IUnknown> m_stored;
    unsigned int      m_store_count = 0;
};
OBJECT_ENTRY_AUTO(CLSID_Registry, RegistryImpl)
