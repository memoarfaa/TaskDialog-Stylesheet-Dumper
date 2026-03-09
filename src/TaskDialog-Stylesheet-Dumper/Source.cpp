#include <windows.h>
#include <shlwapi.h>
#include <xmllite.h>
#include <uxtheme.h>
#include <vssym32.h>          // for TMT_* constants
#include <commdlg.h>
#include <string>
#include <map>
#include <vector>
#include <sstream>
#include <iomanip>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "xmllite.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

// ----------------------------------------------------------------------
// Data structures and global variables
// ----------------------------------------------------------------------
struct ThemeFuncCall {
    std::wstring elementId;
    std::wstring attribute;
    std::wstring functionCall;
    std::wstring funcName;
    std::vector<std::wstring> args;
};

std::vector<ThemeFuncCall> g_themeCalls;
std::wstring g_outputText;
HWND g_hEdit = nullptr;
HWND g_hSaveButton = nullptr;

// ----------------------------------------------------------------------
// Helper to append formatted wide text
// ----------------------------------------------------------------------
void AppendToOutput(const wchar_t* format, ...) {
    wchar_t buffer[4096];
    va_list args;
    va_start(args, format);
    vswprintf_s(buffer, format, args);
    va_end(args);
    g_outputText += buffer;
}

// ----------------------------------------------------------------------
// Parse a function call like "gtf(TaskDialogStyle, 2, 0)"
// ----------------------------------------------------------------------
void ParseFunctionCall(const std::wstring& str, std::wstring& funcName, std::vector<std::wstring>& args) {
    size_t openParen = str.find(L'(');
    if (openParen == std::wstring::npos) return;
    funcName = str.substr(0, openParen);
    size_t closeParen = str.rfind(L')');
    if (closeParen == std::wstring::npos) return;
    std::wstring argsStr = str.substr(openParen + 1, closeParen - openParen - 1);
    size_t start = 0;
    while (start < argsStr.length()) {
        size_t comma = argsStr.find(L',', start);
        if (comma == std::wstring::npos) comma = argsStr.length();
        std::wstring arg = argsStr.substr(start, comma - start);
        // trim spaces
        size_t first = arg.find_first_not_of(L' ');
        if (first != std::wstring::npos) arg = arg.substr(first);
        size_t last = arg.find_last_not_of(L' ');
        if (last != std::wstring::npos) arg = arg.substr(0, last + 1);
        args.push_back(arg);
        start = comma + 1;
    }
}

// ----------------------------------------------------------------------
// Recursive XML traversal to collect theme function calls
// ----------------------------------------------------------------------
void CollectThemeCalls(IXmlReader* pReader, int depth, const std::wstring& currentElementId) {
    XmlNodeType nodeType;
    const wchar_t* localName;

    while (S_OK == pReader->Read(&nodeType)) {
        switch (nodeType) {
        case XmlNodeType_Element:
        {
            pReader->GetLocalName(&localName, nullptr);
            std::wstring elemId = currentElementId;
            if (pReader->MoveToFirstAttribute() == S_OK) {
                do {
                    const wchar_t* attrName, * attrValue;
                    pReader->GetLocalName(&attrName, nullptr);
                    pReader->GetValue(&attrValue, nullptr);
                    if (wcscmp(attrName, L"id") == 0) {
                        elemId = attrValue;
                    }
                    std::wstring attrVal = attrValue;
                    if (attrVal.find(L"gtf(") != std::wstring::npos ||
                        attrVal.find(L"gtc(") != std::wstring::npos ||
                        attrVal.find(L"gtmar(") != std::wstring::npos ||
                        attrVal.find(L"gtmet(") != std::wstring::npos ||
                        attrVal.find(L"dtb(") != std::wstring::npos) {
                        ThemeFuncCall call;
                        call.elementId = elemId;
                        call.attribute = attrName;
                        call.functionCall = attrVal;
                        ParseFunctionCall(attrVal, call.funcName, call.args);
                        g_themeCalls.push_back(call);
                    }
                } while (pReader->MoveToNextAttribute() == S_OK);
                pReader->MoveToElement();
            }
            if (!pReader->IsEmptyElement()) {
                CollectThemeCalls(pReader, depth + 1, elemId);
            }
            break;
        }
        case XmlNodeType_EndElement:
            return;
        default:
            break;
        }
    }
}

// ----------------------------------------------------------------------
// Evaluate a single theme function call using uxtheme
// ----------------------------------------------------------------------
std::wstring EvaluateThemeCall(const ThemeFuncCall& call, HTHEME hThemeTD, HTHEME hThemeTDStyle) {
    std::wstring result;
    HTHEME hTheme = nullptr;

    // Choose the correct theme handle based on the first argument (class name)
    if (call.args.size() >= 1) {
        if (call.args[0] == L"TaskDialog")
            hTheme = hThemeTD;
        else if (call.args[0] == L"TaskDialogStyle")
            hTheme = hThemeTDStyle;
    }

    if (!hTheme) {
        return L"[No theme handle]";
    }

    if (call.funcName == L"gtf") {
        // gtf(class, part, state, prop) - but usually 3 args? Actually gtf(part, state, prop) from XML: gtf(TaskDialogStyle, 2, 0)
        // The class is part of the first arg. So args[0]=class, args[1]=part, args[2]=state, args[3]=prop (optional)
        if (call.args.size() >= 3) {
            int part = _wtoi(call.args[1].c_str());
            int state = _wtoi(call.args[2].c_str());
            int prop = (call.args.size() >= 4) ? _wtoi(call.args[3].c_str()) : TMT_FONT;
            LOGFONTW lf;
            HRESULT hr = GetThemeFont(hTheme, nullptr, part, state, prop, &lf);
            if (SUCCEEDED(hr)) {
                wchar_t buf[256];
                swprintf_s(buf, L"%ls (size %d)", lf.lfFaceName, lf.lfHeight);
                result = buf;
            }
            else {
                result = L"[GetThemeFont failed]";
            }
        }
        else {
            result = L"[Invalid gtf args]";
        }
    }
    else if (call.funcName == L"gtc") {
        // gtc(class, part, state, prop, fallback)
        if (call.args.size() >= 4) {
            int part = _wtoi(call.args[1].c_str());
            int state = _wtoi(call.args[2].c_str());
            int prop = _wtoi(call.args[3].c_str());
            COLORREF color;
            HRESULT hr = GetThemeColor(hTheme, part, state, prop, &color);
            if (SUCCEEDED(hr)) {
                wchar_t buf[64];
                swprintf_s(buf, L"RGB(%d,%d,%d)", GetRValue(color), GetGValue(color), GetBValue(color));
                result = buf;
            }
            else {
                // fallback is a system color index (e.g., 3803 = COLOR_WINDOWTEXT)
                if (call.args.size() >= 5) {
                    int sysColor = _wtoi(call.args[4].c_str());
                    color = GetSysColor(sysColor);
                    wchar_t buf[64];
                    swprintf_s(buf, L"RGB(%d,%d,%d) [fallback]", GetRValue(color), GetGValue(color), GetBValue(color));
                    result = buf;
                }
                else {
                    result = L"[GetThemeColor failed]";
                }
            }
        }
        else {
            result = L"[Invalid gtc args]";
        }
    }
    else if (call.funcName == L"gtmar") {
        // gtmar(class, part, state, prop, fallback)
        if (call.args.size() >= 4) {
            int part = _wtoi(call.args[1].c_str());
            int state = _wtoi(call.args[2].c_str());
            int prop = _wtoi(call.args[3].c_str());
            MARGINS margins;
            HRESULT hr = GetThemeMargins(hTheme, nullptr, part, state, prop, nullptr, &margins);
            if (SUCCEEDED(hr)) {
                wchar_t buf[128];
                swprintf_s(buf, L"left=%d, right=%d, top=%d, bottom=%d",
                    margins.cxLeftWidth, margins.cxRightWidth,
                    margins.cyTopHeight, margins.cyBottomHeight);
                result = buf;
            }
            else {
                result = L"[GetThemeMargins failed]";
            }
        }
        else {
            result = L"[Invalid gtmar args]";
        }
    }
    else if (call.funcName == L"gtmet") {
        // gtmet(class, part, state, prop, fallback)
        if (call.args.size() >= 4) {
            int part = _wtoi(call.args[1].c_str());
            int state = _wtoi(call.args[2].c_str());
            int prop = _wtoi(call.args[3].c_str()); // probably TMT_HEIGHT or similar
            int metric;
            // Try to get as a metric (TMT_HEIGHT is 230? Actually TMT_HEIGHT is 230, but that's a SIZE)
            // Better: use GetThemePartSize to get the default size and extract height.
            SIZE size;
            HDC hdc = GetDC(nullptr);
            HRESULT hr = GetThemePartSize(hTheme, hdc, part, state, nullptr, TS_TRUE, &size);
            ReleaseDC(nullptr, hdc);
            if (SUCCEEDED(hr)) {
                wchar_t buf[64];
                swprintf_s(buf, L"%d pixels (height)", size.cy);
                result = buf;
            }
            else {
                result = L"[GetThemePartSize failed]";
            }
        }
        else {
            result = L"[Invalid gtmet args]";
        }
    }
    else if (call.funcName == L"dtb") {
        // dtb(class, part, state)
        result = L"Uses theme background (DrawThemeBackground)";
    }
    else {
        result = L"[Unknown function]";
    }
    return result;
}

// ----------------------------------------------------------------------
// Load resource, parse, and fill output with XML + theme evaluations
// ----------------------------------------------------------------------
HRESULT DumpTaskDialogStylesheet() {
    HRESULT hr = S_OK;
    HMODULE hMod = LoadLibraryW(L"comctl32.dll");
    if (!hMod) return HRESULT_FROM_WIN32(GetLastError());

    wchar_t path[MAX_PATH];
    GetModuleFileNameW(hMod, path, MAX_PATH);
    AppendToOutput(L"Loaded comctl32.dll from: %s\n\n", path);

    HRSRC hRes = FindResourceW(hMod, MAKEINTRESOURCEW(4255), L"UIFILE");
    if (!hRes) {
        hr = HRESULT_FROM_WIN32(GetLastError());
        FreeLibrary(hMod);
        return hr;
    }

    HGLOBAL hGlobal = LoadResource(hMod, hRes);
    if (!hGlobal) {
        hr = HRESULT_FROM_WIN32(GetLastError());
        FreeLibrary(hMod);
        return hr;
    }

    DWORD size = SizeofResource(hMod, hRes);
    const char* data = (const char*)LockResource(hGlobal);
    if (!data) {
        hr = E_FAIL;
        FreeLibrary(hMod);
        return hr;
    }

    IStream* pStream = SHCreateMemStream((const BYTE*)data, size);
    if (!pStream) {
        hr = E_OUTOFMEMORY;
        FreeLibrary(hMod);
        return hr;
    }

    IXmlReader* pReader = nullptr;
    hr = CreateXmlReader(__uuidof(IXmlReader), (void**)&pReader, nullptr);
    if (FAILED(hr)) {
        pStream->Release();
        FreeLibrary(hMod);
        return hr;
    }

    hr = pReader->SetProperty(XmlReaderProperty_DtdProcessing, DtdProcessing_Prohibit);
    hr = pReader->SetInput(pStream);

    // ------------------------------------------------------------------
    // First pass: extract the <style resid="TaskDialog"> block
    // ------------------------------------------------------------------
    AppendToOutput(L"=== Extracted TaskDialog Stylesheet ===\n\n");

    XmlNodeType nodeType;
    const wchar_t* localName;
    bool inStylesheets = false;
    bool inTargetStyle = false;
    int depth = 0;

    while (S_OK == (hr = pReader->Read(&nodeType))) {
        switch (nodeType) {
        case XmlNodeType_Element:
        {
            pReader->GetLocalName(&localName, nullptr);
            if (!inStylesheets && wcscmp(localName, L"stylesheets") == 0) {
                inStylesheets = true;
                depth = 1;
                AppendToOutput(L"<stylesheets>\n");
            }
            else if (inStylesheets && wcscmp(localName, L"style") == 0) {
                const wchar_t* resid = nullptr;
                hr = pReader->MoveToFirstAttribute();
                while (hr == S_OK) {
                    const wchar_t* attrName, * attrValue;
                    pReader->GetLocalName(&attrName, nullptr);
                    pReader->GetValue(&attrValue, nullptr);
                    if (wcscmp(attrName, L"resid") == 0) {
                        resid = attrValue;
                        break;
                    }
                    hr = pReader->MoveToNextAttribute();
                }
                pReader->MoveToElement();

                if (resid && wcscmp(resid, L"TaskDialog") == 0) {
                    inTargetStyle = true;
                    for (int i = 0; i < depth; i++) AppendToOutput(L"  ");
                    AppendToOutput(L"<style resid=\"TaskDialog\">\n");
                }
                depth++;
            }
            else if (inTargetStyle) {
                for (int i = 0; i < depth; i++) AppendToOutput(L"  ");
                AppendToOutput(L"<%s", localName);
                hr = pReader->MoveToFirstAttribute();
                while (hr == S_OK) {
                    const wchar_t* attrName, * attrValue;
                    pReader->GetLocalName(&attrName, nullptr);
                    pReader->GetValue(&attrValue, nullptr);
                    AppendToOutput(L" %ls=\"%ls\"", attrName, attrValue);
                    hr = pReader->MoveToNextAttribute();
                }
                pReader->MoveToElement();
                if (pReader->IsEmptyElement())
                    AppendToOutput(L"/>\n");
                else
                    AppendToOutput(L">\n");
                depth++;
            }
            else {
                depth++;
            }
            break;
        }
        case XmlNodeType_EndElement:
        {
            pReader->GetLocalName(&localName, nullptr);
            depth--;
            if (inTargetStyle) {
                if (wcscmp(localName, L"style") == 0) {
                    for (int i = 0; i < depth; i++) AppendToOutput(L"  ");
                    AppendToOutput(L"</style>\n");
                    inTargetStyle = false;
                }
                else {
                    for (int i = 0; i < depth; i++) AppendToOutput(L"  ");
                    AppendToOutput(L"</%ls>\n", localName);
                }
            }
            if (inStylesheets && wcscmp(localName, L"stylesheets") == 0) {
                AppendToOutput(L"</stylesheets>\n");
                inStylesheets = false;
            }
            break;
        }
        case XmlNodeType_Text:
            if (inTargetStyle) {
                LPCWSTR value;
                pReader->GetValue(&value, nullptr);
                if (value && *value) {
                    for (int i = 0; i < depth; i++) AppendToOutput(L"  ");
                    AppendToOutput(L"%ls\n", value);
                }
            }
            break;
        default:
            break;
        }
    }

    pReader->Release();
    pStream->Release();

    // ------------------------------------------------------------------
    // Second pass: collect theme function calls
    // ------------------------------------------------------------------
    pStream = SHCreateMemStream((const BYTE*)data, size);
    if (!pStream) {
        FreeLibrary(hMod);
        return E_OUTOFMEMORY;
    }
    hr = CreateXmlReader(__uuidof(IXmlReader), (void**)&pReader, nullptr);
    if (FAILED(hr)) {
        pStream->Release();
        FreeLibrary(hMod);
        return hr;
    }
    pReader->SetProperty(XmlReaderProperty_DtdProcessing, DtdProcessing_Prohibit);
    pReader->SetInput(pStream);

    g_themeCalls.clear();
    CollectThemeCalls(pReader, 0, L"");

    pReader->Release();
    pStream->Release();

    // ------------------------------------------------------------------
    // Output explanatory table
    // ------------------------------------------------------------------
    AppendToOutput(L"\n\n=== How the XML Uses comctl32.dll Theme Functions ===\n\n");
    AppendToOutput(L"| Function | Meaning | Example in XML |\n");
    AppendToOutput(L"|----------|---------|----------------|\n");

    auto addRow = [&](const wchar_t* func, const wchar_t* meaning, const wchar_t* example) {
        AppendToOutput(L"| **`%s`** | %s | `%s` |\n", func, meaning, example);
        };

    addRow(L"gtf(part, state, prop)",
        L"Get Theme Font – retrieves a font from the current theme.",
        L"font=\"gtf(TaskDialogStyle, 2, 0)\"");
    addRow(L"gtc(part, state, prop, fallback)",
        L"Get Theme Color – retrieves a color; fallback is a system color index.",
        L"foreground=\"gtc(TaskDialogStyle, 4, 0, 3803)\"");
    addRow(L"gtmar(part, state, prop, fallback)",
        L"Get Theme Margins – retrieves margin or padding values.",
        L"padding=\"gtmar(TaskDialog, 1, 0, 3602)\"");
    addRow(L"gtmet(part, state, prop, fallback)",
        L"Get Theme Metric – retrieves an integer metric (height, width).",
        L"height=\"gtmet(TaskDialog, 19, 0, 2417)\"");
    addRow(L"dtb(part, state)",
        L"Draw Theme Background – renders using the theme part/state.",
        L"background=\"dtb(TaskDialog, 1, 0)\"");

    AppendToOutput(L"\n**Note:** Many theme functions are wrapped in `themeable()` (e.g., `background=\"themeable(dtb(...), threedface)\"`) to provide a fallback when visual styles are disabled.\n");

    // ------------------------------------------------------------------
    // Evaluate theme calls using uxtheme
    // ------------------------------------------------------------------
    AppendToOutput(L"\n\n=== Actual Theme Values (Current Visual Style) ===\n\n");

    HTHEME hThemeTaskDialog = OpenThemeData(nullptr, L"TaskDialog");
    HTHEME hThemeTaskDialogStyle = OpenThemeData(nullptr, L"TaskDialogStyle");

    if (!hThemeTaskDialog && !hThemeTaskDialogStyle) {
        AppendToOutput(L"Could not open theme data. Are visual styles enabled?\n");
    }
    else {
        for (const auto& call : g_themeCalls) {
            std::wstring eval = EvaluateThemeCall(call, hThemeTaskDialog, hThemeTaskDialogStyle);
            AppendToOutput(L"  [%ls] %ls = \"%ls\" → %ls\n",
                call.elementId.c_str(),
                call.attribute.c_str(),
                call.functionCall.c_str(),
                eval.c_str());
        }

        if (hThemeTaskDialog) CloseThemeData(hThemeTaskDialog);
        if (hThemeTaskDialogStyle) CloseThemeData(hThemeTaskDialogStyle);
    }

    FreeLibrary(hMod);
    return S_OK;
}

// ----------------------------------------------------------------------
// Save the current content to an XML file
// ----------------------------------------------------------------------
void SaveToFile(HWND hWnd) {
    OPENFILENAMEW ofn = { sizeof(ofn) };
    wchar_t szFile[260] = { 0 };

    ofn.hwndOwner = hWnd;
    ofn.lpstrFilter = L"XML Files\0*.xml\0All Files\0*.*\0";
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = ARRAYSIZE(szFile);
    ofn.lpstrDefExt = L"xml";
    ofn.Flags = OFN_OVERWRITEPROMPT;

    if (GetSaveFileNameW(&ofn)) {
        HANDLE hFile = CreateFileW(szFile, GENERIC_WRITE, 0, nullptr,
            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (hFile != INVALID_HANDLE_VALUE) {
            DWORD written;
            const WORD bom = 0xFEFF;
            WriteFile(hFile, &bom, sizeof(bom), &written, nullptr);
            WriteFile(hFile, g_outputText.c_str(), (DWORD)(g_outputText.size() * sizeof(wchar_t)), &written, nullptr);
            CloseHandle(hFile);
            MessageBoxW(hWnd, L"File saved successfully.", L"Success", MB_OK | MB_ICONINFORMATION);
        }
        else {
            MessageBoxW(hWnd, L"Failed to create file.", L"Error", MB_OK | MB_ICONERROR);
        }
    }
}

// ----------------------------------------------------------------------
// Window procedure
// ----------------------------------------------------------------------
LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE:
    {
        g_hEdit = CreateWindowW(L"EDIT", nullptr,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
            10, 10, 600, 300,
            hWnd, nullptr, GetModuleHandle(nullptr), nullptr);
        g_hSaveButton = CreateWindowW(L"BUTTON", L"Save as XML",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            250, 320, 120, 30,
            hWnd, (HMENU)1, GetModuleHandle(nullptr), nullptr);
        HRESULT hr = DumpTaskDialogStylesheet();
        if (FAILED(hr)) {
            wchar_t errMsg[100];
            swprintf_s(errMsg, L"Failed to load stylesheet. HRESULT: 0x%08X", hr);
            SetWindowTextW(g_hEdit, errMsg);
        }
        else {
            SetWindowTextW(g_hEdit, g_outputText.c_str());
        }
    }
    break;

    case WM_SIZE:
    {
        RECT rc;
        GetClientRect(hWnd, &rc);
        SetWindowPos(g_hEdit, nullptr, 10, 10, rc.right - 20, rc.bottom - 60, SWP_NOZORDER);
        SetWindowPos(g_hSaveButton, nullptr, (rc.right - 120) / 2, rc.bottom - 40, 120, 30, SWP_NOZORDER);
    }
    break;

    case WM_COMMAND:
        if (LOWORD(wParam) == 1) {
            SaveToFile(hWnd);
        }
        break;

    case WM_DESTROY:
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProcW(hWnd, msg, wParam, lParam);
    }
    return 0;
}

// ----------------------------------------------------------------------
// Entry point
// ----------------------------------------------------------------------
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int nCmdShow) {
    WNDCLASSW wc = { 0 };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = L"TaskDialogDumperClass";
    RegisterClassW(&wc);

    HWND hWnd = CreateWindowW(wc.lpszClassName, L"TaskDialog Stylesheet Dumper",
        WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
        CW_USEDEFAULT, CW_USEDEFAULT, 640, 400,
        nullptr, nullptr, hInstance, nullptr);
    if (!hWnd) return 0;

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return (int)msg.wParam;
}