#include "NonWindows.hpp"


__attribute__((visibility("default")))
BSTR SysAllocStringLen (const wchar_t* str, unsigned int len) {
    if (len > (UINT32_MAX - sizeof(wchar_t))/sizeof(wchar_t))
        return nullptr; // byte length doesn't fit in the 32bit prefix

    const size_t byte_len = static_cast<size_t>(len)*sizeof(wchar_t);

    // byte length prefix, followed by the characters and a zero terminator
    auto* ptr = static_cast<unsigned char*>(malloc(sizeof(uint32_t) + byte_len + sizeof(wchar_t)));
    if (!ptr)
        return nullptr;

    *reinterpret_cast<uint32_t*>(ptr) = static_cast<uint32_t>(byte_len);
    BSTR result = reinterpret_cast<BSTR>(ptr + sizeof(uint32_t));
    if (str)
        memcpy(result, str, byte_len); // might contain embedded zeros
    else
        memset(result, 0, byte_len);
    result[len] = L'\0';
    return result;
}

__attribute__((visibility("default")))
BSTR SysAllocString (const wchar_t* str) {
    if (!str)
        return nullptr;

    return SysAllocStringLen(str, static_cast<unsigned int>(wcslen(str)));
}

__attribute__((visibility("default")))
void SysFreeString (BSTR str) {
    if (!str)
        return;

    free(reinterpret_cast<unsigned char*>(str) - sizeof(uint32_t));
}

__attribute__((visibility("default")))
unsigned int SysStringByteLen (BSTR str) {
    if (!str)
        return 0;

    return *reinterpret_cast<uint32_t*>(reinterpret_cast<unsigned char*>(str) - sizeof(uint32_t));
}

__attribute__((visibility("default")))
unsigned int SysStringLen (BSTR str) {
    return SysStringByteLen(str)/sizeof(wchar_t);
}


__attribute__((visibility("default")))
void* CoTaskMemAlloc (size_t size) {
    return malloc(size);
}

__attribute__((visibility("default")))
void* CoTaskMemRealloc (void* ptr, size_t size) {
    return realloc(ptr, size);
}

__attribute__((visibility("default")))
void CoTaskMemFree (void* ptr) {
    free(ptr);
}


/** Resolve COM class CLSID based on "[<Program>.]<Component>[.<Version>]" ProgID string. */
__attribute__((visibility("default")))
HRESULT CLSIDFromProgID (const wchar_t* ProgID, /*out*/GUID* clsid) {
    if (!ProgID || !clsid)
        return E_POINTER;
    
    std::wstring class_name = ProgID;

    // remove "<Program>." prefix and ".<Version>" suffix if present
    size_t idx1 = class_name.find(L'.');
    if (idx1 != std::wstring::npos) {
        std::wstring suffix = class_name.substr(idx1+1); // "<Component>.<Version>" or "<Version>"
        size_t idx2 = suffix.find(L'.');

        if (idx2 != std::wstring::npos) {
            // input contain two '.'s, keep center part
            class_name = suffix.substr(0,idx2); // "<Component>"
        } else {
            // input contain one '.'. check if suffix is a number
            auto version = wcstol(suffix.c_str(), nullptr, /*base*/10);
            if (version != 0) {
                class_name = class_name.substr(0, idx1);
            } else {
                class_name = suffix;
            }
        }
    }

    for (size_t i = 0; i < IUnknownFactory::Factories().size(); i++) {
        const auto & elm = IUnknownFactory::Factories()[i];
        if (elm.name == ATL::CComBSTR(class_name.c_str())) {
            *clsid = elm.clsid;
            return S_OK;
        }
    }

    std::wcerr << L"CLSIDFromProgID error: Unknown class " << class_name << std::endl;
    return CO_E_CLASSSTRING;
}

/** Create COM class based on CLSID. */
__attribute__((visibility("default")))
HRESULT CoCreateInstance (const GUID& clsid, IUnknown* outer, DWORD context, const GUID& iid, /*out*/void** result) {
    (void)context;

    if (!result)
        return E_POINTER;
    *result = nullptr;

    if (outer && !(iid == IID_IUnknown))
        return E_INVALIDARG;

    for (size_t i = 0; i < IUnknownFactory::Factories().size(); i++) {
        const auto & elm = IUnknownFactory::Factories()[i];
        if (elm.clsid == clsid) {
            IUnknown* obj = nullptr;
            HRESULT hr = elm.factory(outer, &obj); // RefCount=1
            assert(hr == S_OK);
            if (FAILED(hr))
                return hr;

            // cast to requested interface
            assert(obj);
            hr = obj->QueryInterface(iid, result);
            obj->Release();
            obj = nullptr;
            return hr;
        }
    }

    char guid_str[39] = {};
    snprintf(guid_str, sizeof(guid_str), "{%08x-%04x-%04x-%02x%02x-%02x%02x%02x%02x%02x%02x}",
        clsid.Data1, clsid.Data2, clsid.Data3,
        clsid.Data4[0], clsid.Data4[1], clsid.Data4[2], clsid.Data4[3],
        clsid.Data4[4], clsid.Data4[5], clsid.Data4[6], clsid.Data4[7]);

    std::cerr << "CoCreateInstance error: Unknown clsid " << guid_str << std::endl;
    return REGDB_E_CLASSNOTREG;
}


__attribute__((visibility("default")))
Buffer<IUnknownFactory::Entry> & IUnknownFactory::Factories () {
    static Buffer<Entry> s_factory;
    return s_factory;
}


/** Hidden bytes in front of the SAFEARRAY descriptor, holding the element VARTYPE, like on Windows. */
static constexpr size_t SAFEARRAY_HIDDEN_SIZE = 16;

__attribute__((visibility("default")))
SAFEARRAY* InternalSafeArrayCreate (VARTYPE vt, ULONG elm_size, ULONG count) {
    auto* ptr = static_cast<unsigned char*>(calloc(1, SAFEARRAY_HIDDEN_SIZE + sizeof(SAFEARRAY)));
    if (!ptr)
        return nullptr;

    auto* psa = reinterpret_cast<SAFEARRAY*>(ptr + SAFEARRAY_HIDDEN_SIZE);
    psa->cDims = 1;
    psa->cbElements = elm_size;
    psa->rgsabound[0].cElements = count;
    if (vt != VT_EMPTY) {
        psa->fFeatures = FADF_HAVEVARTYPE | ((vt == VT_BSTR) ? FADF_BSTR : (vt == VT_UNKNOWN) ? FADF_UNKNOWN : 0);
        reinterpret_cast<uint32_t*>(psa)[-1] = vt; // stored in the 4 bytes before the descriptor
    }

    psa->pvData = calloc(count ? count : 1, elm_size); // zero-initialized
    if (!psa->pvData) {
        free(ptr);
        return nullptr;
    }
    return psa;
}

__attribute__((visibility("default")))
SAFEARRAY* SafeArrayCreate (VARTYPE vt, UINT cDims, SAFEARRAYBOUND* rgsabound) {
    if ((cDims != 1) || !rgsabound)
        return nullptr; // only one-dimensional arrays are supported

    const ULONG elm_size = InternalVartypeSize(vt);
    if (!elm_size)
        return nullptr; // unsupported element type

    SAFEARRAY* psa = InternalSafeArrayCreate(vt, elm_size, rgsabound->cElements);
    if (psa)
        psa->rgsabound[0].lLbound = rgsabound->lLbound;
    return psa;
}

__attribute__((visibility("default")))
SAFEARRAY* SafeArrayCreateVector (VARTYPE vt, LONG lLbound, ULONG cElements) {
    SAFEARRAYBOUND bound = {cElements, lLbound};
    return SafeArrayCreate(vt, 1, &bound);
}

__attribute__((visibility("default")))
HRESULT SafeArrayDestroy (SAFEARRAY* psa) {
    if (!psa)
        return S_OK;
    if (psa->cLocks)
        return DISP_E_ARRAYISLOCKED;

    const ULONG count = psa->rgsabound[0].cElements;
    if (psa->fFeatures & FADF_BSTR) {
        for (ULONG i = 0; i < count; i++)
            SysFreeString(static_cast<BSTR*>(psa->pvData)[i]);
    } else if (psa->fFeatures & FADF_UNKNOWN) {
        for (ULONG i = 0; i < count; i++) {
            IUnknown* obj = static_cast<IUnknown**>(psa->pvData)[i];
            if (obj)
                obj->Release();
        }
    }
    free(psa->pvData);
    free(reinterpret_cast<unsigned char*>(psa) - SAFEARRAY_HIDDEN_SIZE);
    return S_OK;
}

__attribute__((visibility("default")))
HRESULT SafeArrayGetVartype (SAFEARRAY* psa, /*out*/VARTYPE* pvt) {
    if (!psa || !pvt)
        return E_INVALIDARG;
    if (!(psa->fFeatures & FADF_HAVEVARTYPE))
        return E_INVALIDARG;

    *pvt = static_cast<VARTYPE>(reinterpret_cast<uint32_t*>(psa)[-1]); // stored in the 4 bytes before the descriptor
    return S_OK;
}

__attribute__((visibility("default")))
UINT SafeArrayGetDim (SAFEARRAY* psa) {
    return psa ? psa->cDims : 0;
}

__attribute__((visibility("default")))
UINT SafeArrayGetElemsize (SAFEARRAY* psa) {
    return psa ? psa->cbElements : 0;
}

__attribute__((visibility("default")))
HRESULT SafeArrayGetLBound (SAFEARRAY* psa, UINT nDim, /*out*/LONG* plLbound) {
    if (!psa || !plLbound)
        return E_INVALIDARG;
    if (nDim != 1)
        return DISP_E_BADINDEX;

    *plLbound = psa->rgsabound[0].lLbound;
    return S_OK;
}

__attribute__((visibility("default")))
HRESULT SafeArrayGetUBound (SAFEARRAY* psa, UINT nDim, /*out*/LONG* plUbound) {
    if (!psa || !plUbound)
        return E_INVALIDARG;
    if (nDim != 1)
        return DISP_E_BADINDEX;

    *plUbound = psa->rgsabound[0].lLbound + static_cast<LONG>(psa->rgsabound[0].cElements) - 1;
    return S_OK;
}

__attribute__((visibility("default")))
HRESULT SafeArrayAccessData (SAFEARRAY* psa, /*out*/void** ppvData) {
    if (!psa || !ppvData)
        return E_INVALIDARG;

    psa->cLocks++;
    *ppvData = psa->pvData;
    return S_OK;
}

__attribute__((visibility("default")))
HRESULT SafeArrayUnaccessData (SAFEARRAY* psa) {
    if (!psa)
        return E_INVALIDARG;
    if (psa->cLocks == 0)
        return E_UNEXPECTED;

    psa->cLocks--;
    return S_OK;
}
