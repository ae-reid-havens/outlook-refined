// ==WindhawkMod==
// @id              outlook-refined
// @name            Outlook Refined
// @description     Restyles new Outlook: favorites groups, per-account compose buttons, a resizable folder pane, an agenda with a month grid, the Other inbox under Focused, a tidier calendar and reading pane
// @version         0.2
// @author          Analytic Endeavors (Reid Havens)
// @github          https://github.com/ae-reid-havens
// @include         olk.exe
// ==/WindhawkMod==

// ==WindhawkModReadme==
/*
# Outlook Refined

Restyles the new Outlook for Windows (olk.exe) from inside its own page. The mod
adds one script to each Outlook window when the window is created. It reads
nothing from disk and opens no ports. The agenda, month grid and Other inbox
card ask Outlook's own mail service for your calendar and inbox, from inside
its page and with the session Outlook is already signed in to; nothing is
stored, logged or sent anywhere else.

Restart Outlook after enabling the mod or changing a setting.

**Folder pane:** Favorites in groups that fold (right-click a group header for
its menu, Shift+right-click a favorite to rename or move it), per-account
compose buttons, a resizable pane, account trees under one foldable heading,
accounts reordered by dragging. Edits made inside Outlook are kept per PC in
Outlook's own storage.

**Agenda:** today and the next days from every account's calendar at the
bottom of the folder pane, with Outlook's event icons and a join button for
online meetings, under a month grid shaded by booked hours. **Other inbox:**
optionally, the newest Other conversations below the Focused list.

**Calendar:** one row per account, compact time zones, start and end times on
event tiles, overlapping events cascaded or merged.

**Reading pane:** To and Cc on one line, a compact blocked-content bar, a
readable body width. **Ribbon:** the group-name row under the buttons goes.
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- favoritesGroups: true
  $name: Favorites groups
  $description: Show favorites in groups that fold. Starts with an Outgoing group (Drafts, Sent, Outbox) and a Junk group
- hiddenFolders: [""]
  $name: Hidden favorites
  $description: Favorite folder names to hide, as Outlook shows them (for example "Archive")
- composeButtons:
  - - account: ""
      $name: Account address
    - label: ""
      $name: Button label
    - color: "#0f6cbd"
      $name: Button color
  $name: Compose buttons
  $description: One button per account at the top of Favorites; a new email starts from that account
- accountLabels:
  - - account: ""
      $name: Account address
    - label: ""
      $name: Short label
  $name: Account labels
  $description: Short names for accounts in the folder pane, pattern groups and the calendar list
- primaryAccount: ""
  $name: Primary account address
  $description: Leave empty to detect it. Outlook names the primary account's folders without an address suffix
- accountsHeading: Accounts
  $name: Accounts heading
  $description: Heading that folds every account tree. Empty for none
- hideAccountTrees: false
  $name: Hide account trees
- hideAddAccount: false
  $name: Hide "Add account"
- resizablePane: true
  $name: Resizable folder pane
  $description: Drag the pane's right edge; double-click it to reset
- paneWidth: 0
  $name: Folder pane width (px)
  $description: 0 keeps Outlook's own width. A width set by dragging the pane's edge wins; double-click the edge to return to this one
- agenda: true
  $name: Agenda under the folder pane
  $description: Today and the next days from every account's calendar, with a join button for online meetings. Read through the session Outlook is signed in to; nothing is stored or sent anywhere else
- agendaDays: 3
  $name: Agenda days
  $description: Days the agenda lists, today included (1 to 14)
- monthHeatMap: true
  $name: Month grid above the agenda
  $description: Each day shaded by its booked hours; the arrows browse other months
- otherCard: false
  $name: Other inbox under Focused
  $description: While an inbox shows its Focused tab, its newest Other conversations sit below the list. A button by the tabs switches back to Outlook's own tabs
- calendarRowPerAccount: true
  $name: Calendar, one row per account
- calendarPaneWidth: 180
  $name: Calendar pane width (px)
  $description: 0 keeps Outlook's own width
- compactTimeZones: true
  $name: Calendar, compact time zones
- tileTimes: true
  $name: Calendar, times on event tiles
- overlapLayout: merge
  $name: Calendar, overlapping events
  $options:
  - none: Outlook's layout
  - cascade: Cascade
  - merge: Cascade, and merge larger overlaps into one tile
- mergeAt: 3
  $name: Merge at this many overlapping events
- oneLineRecipients: true
  $name: Reading pane, To and Cc on one line
- blockedContentChip: true
  $name: Reading pane, compact blocked-content bar
- readingWidth: 90ch
  $name: Reading pane, body width
  $description: A CSS width such as 90ch or 800px. Empty for Outlook's own
- hideRibbonGroupLabels: true
  $name: Ribbon, hide group names
*/
// ==/WindhawkModSettings==

#include <unknwn.h>
#include <atomic>
#include <string>

// Interface ids and vtable slots from WebView2.h (Microsoft.Web.WebView2 SDK
// 1.0.3179.45, BSD licence). COM vtables are a stable ABI, so slots never move.
constexpr GUID kIID_IUnknown = {0x00000000, 0x0000, 0x0000, {0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};
constexpr GUID kIID_EnvCompleted = {0x4e8a3389, 0xc9d8, 0x4bd2, {0xb6, 0xb5, 0x12, 0x4f, 0xee, 0x6c, 0xc1, 0x4d}};
constexpr GUID kIID_ControllerCompleted = {0x6c4819f3, 0xc9b7, 0x4260, {0x81, 0x27, 0xc9, 0xf5, 0xbd, 0xe7, 0xf6, 0x8c}};
constexpr GUID kIID_CompositionCompleted = {0x02fab84b, 0x1428, 0x4fb7, {0xad, 0x45, 0x1b, 0x2e, 0x64, 0x73, 0x61, 0x84}};
constexpr GUID kIID_ScriptAdded = {0xb99369f3, 0x9b11, 0x47b5, {0xbc, 0x6f, 0x8e, 0x78, 0x95, 0xfc, 0xea, 0x17}};
constexpr GUID kIID_Env3 = {0x80a22ae3, 0xbe7c, 0x4ce2, {0xaf, 0xe1, 0x5a, 0x50, 0x05, 0x6c, 0xde, 0xeb}};
constexpr GUID kIID_Env10 = {0xee0eb9df, 0x6f12, 0x46ce, {0xb5, 0x3f, 0x3f, 0x47, 0xb9, 0xc9, 0x28, 0xe0}};
constexpr GUID kIID_Controller = {0x4d00c0d1, 0x9434, 0x4eb6, {0x80, 0x78, 0x86, 0x97, 0xa5, 0x60, 0x33, 0x4f}};

enum {
    kEnv_CreateController = 3,
    kEnv3_CreateCompositionController = 9,
    kEnv10_CreateControllerWithOptions = 21,
    kEnv10_CreateCompositionControllerWithOptions = 22,
    kController_GetCoreWebView2 = 25,
    kWebView_AddScriptOnDocumentCreated = 27,
    kHandler_Invoke = 3,
};

// The page script (src/page.js), stitched in by build.mjs.
const wchar_t kPageScript[] = LR"olkt(/*@@PAGE@@*/)olkt";

// ==== settings to OLKT_CONFIG ====

// A JSON string literal. U+2028 and U+2029 are escaped too: JSON allows them raw,
// older JavaScript string literals do not.
std::wstring JsonString(PCWSTR text) {
    std::wstring out = L"\"";
    for (const wchar_t* p = text; *p; ++p) {
        wchar_t c = *p;
        if (c == L'"' || c == L'\\') {
            out += L'\\';
            out += c;
        } else if (c < 0x20 || c == 0x2028 || c == 0x2029) {
            wchar_t buf[8];
            swprintf(buf, 8, L"\\u%04x", (unsigned)c);
            out += buf;
        } else {
            out += c;
        }
    }
    return out + L"\"";
}

std::wstring StringSetting(PCWSTR name, ...) {
    va_list args;
    va_start(args, name);
    wchar_t key[128];
    vswprintf(key, 128, name, args);
    va_end(args);
    PCWSTR value = Wh_GetStringSetting(key);
    std::wstring out = value;
    Wh_FreeStringSetting(value);
    return out;
}

void AddField(std::wstring& json, PCWSTR name, const std::wstring& value) {
    if (json.size() > 1) json += L",";
    json += JsonString(name) + L":" + value;
}

void AddBool(std::wstring& json, PCWSTR name) { AddField(json, name, Wh_GetIntSetting(name) ? L"true" : L"false"); }
void AddInt(std::wstring& json, PCWSTR name) { AddField(json, name, std::to_wstring(Wh_GetIntSetting(name))); }
void AddString(std::wstring& json, PCWSTR name) { AddField(json, name, JsonString(StringSetting(name).c_str())); }

// A list setting as a JSON array. `fields` empty: a list of strings; otherwise a
// list of objects with those fields. Entries whose first field is empty are skipped.
void AddList(std::wstring& json, PCWSTR name, std::initializer_list<PCWSTR> fields) {
    std::wstring arr = L"[";
    for (int i = 0; i < 100; ++i) {
        std::wstring item;
        if (fields.size() == 0) {
            std::wstring value = StringSetting(L"%s[%d]", name, i);
            if (value.empty()) continue;
            item = JsonString(value.c_str());
        } else {
            item = L"{";
            bool first = true;
            for (PCWSTR field : fields) {
                std::wstring value = StringSetting(L"%s[%d].%s", name, i, field);
                if (first && value.empty()) break;
                if (!first) item += L",";
                item += JsonString(field) + L":" + JsonString(value.c_str());
                first = false;
            }
            if (first) continue;
            item += L"}";
        }
        if (arr.size() > 1) arr += L",";
        arr += item;
    }
    AddField(json, name, arr + L"]");
}

SRWLOCK g_scriptLock = SRWLOCK_INIT;
std::wstring g_script;

// Wraps the page as `{ const OLKT_CONFIG = {...}; <page> }`. Windows created after a
// settings change get the new settings; open windows keep theirs until a restart.
void BuildScript() {
    std::wstring json = L"{";
    AddBool(json, L"favoritesGroups");
    AddList(json, L"hiddenFolders", {});
    AddList(json, L"composeButtons", {L"account", L"label", L"color"});
    AddList(json, L"accountLabels", {L"account", L"label"});
    AddString(json, L"primaryAccount");
    AddString(json, L"accountsHeading");
    AddBool(json, L"hideAccountTrees");
    AddBool(json, L"hideAddAccount");
    AddBool(json, L"resizablePane");
    AddInt(json, L"paneWidth");
    AddBool(json, L"agenda");
    AddInt(json, L"agendaDays");
    AddBool(json, L"monthHeatMap");
    AddBool(json, L"otherCard");
    AddBool(json, L"calendarRowPerAccount");
    AddInt(json, L"calendarPaneWidth");
    AddBool(json, L"compactTimeZones");
    AddBool(json, L"tileTimes");
    AddString(json, L"overlapLayout");
    AddInt(json, L"mergeAt");
    AddBool(json, L"oneLineRecipients");
    AddBool(json, L"blockedContentChip");
    AddString(json, L"readingWidth");
    AddBool(json, L"hideRibbonGroupLabels");
    json += L"}";
    std::wstring script = L"{\nconst OLKT_CONFIG = " + json + L";\n" + kPageScript + L"\n}";
    AcquireSRWLockExclusive(&g_scriptLock);
    g_script.swap(script);
    ReleaseSRWLockExclusive(&g_scriptLock);
}

// ==== WebView2 hooks ====

void* Slot(IUnknown* obj, int index) {
    return (*reinterpret_cast<void***>(obj))[index];
}

// COM objects of ours that WebView2 still holds. Their code lives in this DLL, so
// Wh_ModUninit waits for them: a call into an unloaded mod would crash Outlook.
std::atomic<int> g_liveObjects{0};

using Invoke_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*, HRESULT, IUnknown*);
HRESULT Forward(IUnknown* handler, HRESULT hr, IUnknown* result) {
    return reinterpret_cast<Invoke_t>(Slot(handler, kHandler_Invoke))(handler, hr, result);
}

// Completion handler for AddScriptToExecuteOnDocumentCreated. Holds back Outlook's
// controller handler until the script is registered, because Outlook navigates from
// that handler and the window's first document must already get the script.
class ScriptAdded final : public IUnknown {
    std::atomic<ULONG> refs_{1};
    std::atomic<bool> forwarded_{false};
    IUnknown* outlook_;
    IUnknown* controller_;

   public:
    ScriptAdded(IUnknown* outlook, IUnknown* controller) : outlook_(outlook), controller_(controller) {
        outlook_->AddRef();
        controller_->AddRef();
        ++g_liveObjects;
    }
    // Hands Outlook its controller, once only.
    void ForwardOnce() {
        if (!forwarded_.exchange(true)) Forward(outlook_, S_OK, controller_);
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** out) override {
        if (IsEqualGUID(iid, kIID_IUnknown) || IsEqualGUID(iid, kIID_ScriptAdded)) {
            *out = this;
            AddRef();
            return S_OK;
        }
        *out = nullptr;
        return E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return ++refs_; }
    ULONG STDMETHODCALLTYPE Release() override {
        ULONG n = --refs_;
        if (!n) {
            ForwardOnce();  // dropped without a completion: Outlook still gets its window
            outlook_->Release();
            controller_->Release();
            delete this;
            --g_liveObjects;
        }
        return n;
    }
    virtual HRESULT STDMETHODCALLTYPE Invoke(HRESULT hr, LPCWSTR id) {
        if (FAILED(hr)) Wh_Log(L"Page script not added: hr=0x%08X", hr);
        ForwardOnce();
        return S_OK;
    }
};

// Stands in for Outlook's completion handler. `handle` gets the result and Outlook's
// handler, and is responsible for passing the result on to it.
using Handle_t = HRESULT (*)(HRESULT hr, IUnknown* result, IUnknown* outlook);

class WrappedHandler final : public IUnknown {
    std::atomic<ULONG> refs_{1};
    GUID iid_;
    IUnknown* inner_;
    Handle_t handle_;

   public:
    WrappedHandler(REFIID iid, IUnknown* inner, Handle_t handle) : iid_(iid), inner_(inner), handle_(handle) {
        inner_->AddRef();
        ++g_liveObjects;
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** out) override {
        if (IsEqualGUID(iid, kIID_IUnknown) || IsEqualGUID(iid, iid_)) {
            *out = this;
            AddRef();
            return S_OK;
        }
        *out = nullptr;
        return E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return ++refs_; }
    ULONG STDMETHODCALLTYPE Release() override {
        ULONG n = --refs_;
        if (!n) {
            inner_->Release();
            delete this;
            --g_liveObjects;
        }
        return n;
    }
    virtual HRESULT STDMETHODCALLTYPE Invoke(HRESULT hr, IUnknown* result) { return handle_(hr, result, inner_); }
};

// Set while one of our controller hooks runs, so an implementation that calls
// another hooked method internally is wrapped only once.
thread_local int g_inCreate = 0;

// Adds the page script to each new WebView, then hands Outlook its controller once
// the script is registered (ScriptAdded). Any failure hands it over at once.
HRESULT OnController(HRESULT result_hr, IUnknown* result, IUnknown* outlook) {
    if (FAILED(result_hr) || !result) return Forward(outlook, result_hr, result);
    IUnknown* controller = nullptr;
    if (FAILED(result->QueryInterface(kIID_Controller, (void**)&controller))) {
        Wh_Log(L"Controller interface missing");
        return Forward(outlook, result_hr, result);
    }
    IUnknown* webview = nullptr;
    using Get_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*, IUnknown**);
    HRESULT hr = reinterpret_cast<Get_t>(Slot(controller, kController_GetCoreWebView2))(controller, &webview);
    controller->Release();
    if (FAILED(hr) || !webview) {
        Wh_Log(L"get_CoreWebView2 failed: 0x%08X", hr);
        return Forward(outlook, result_hr, result);
    }
    AcquireSRWLockShared(&g_scriptLock);
    std::wstring script = g_script;
    ReleaseSRWLockShared(&g_scriptLock);
    auto* done = new ScriptAdded(outlook, result);
    using Add_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*, LPCWSTR, IUnknown*);
    hr = reinterpret_cast<Add_t>(Slot(webview, kWebView_AddScriptOnDocumentCreated))(webview, script.c_str(), done);
    webview->Release();
    Wh_Log(L"New WebView: AddScriptToExecuteOnDocumentCreated hr=0x%08X", hr);
    if (FAILED(hr)) done->ForwardOnce();  // no completion will come
    done->Release();
    return S_OK;
}

using CreateController_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*, HWND, IUnknown*);
using CreateControllerOpt_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*, HWND, IUnknown*, IUnknown*);
CreateController_t CreateController_Original;
CreateController_t CreateComposition_Original;
CreateControllerOpt_t CreateControllerOpt_Original;
CreateControllerOpt_t CreateCompositionOpt_Original;

HRESULT CallWrapped(REFIID iid, IUnknown* handler, auto call) {
    ++g_inCreate;
    IUnknown* h = handler;
    if (h && g_inCreate == 1) {
        h = new WrappedHandler(iid, handler, OnController);
    } else if (h) {
        h->AddRef();
    }
    HRESULT hr = call(h);
    if (h) h->Release();
    --g_inCreate;
    return hr;
}

HRESULT STDMETHODCALLTYPE CreateController_Hook(IUnknown* env, HWND parent, IUnknown* handler) {
    return CallWrapped(kIID_ControllerCompleted, handler,
                       [&](IUnknown* h) { return CreateController_Original(env, parent, h); });
}

HRESULT STDMETHODCALLTYPE CreateComposition_Hook(IUnknown* env, HWND parent, IUnknown* handler) {
    return CallWrapped(kIID_CompositionCompleted, handler,
                       [&](IUnknown* h) { return CreateComposition_Original(env, parent, h); });
}

HRESULT STDMETHODCALLTYPE CreateControllerOpt_Hook(IUnknown* env, HWND parent, IUnknown* opts, IUnknown* handler) {
    return CallWrapped(kIID_ControllerCompleted, handler,
                       [&](IUnknown* h) { return CreateControllerOpt_Original(env, parent, opts, h); });
}

HRESULT STDMETHODCALLTYPE CreateCompositionOpt_Hook(IUnknown* env, HWND parent, IUnknown* opts, IUnknown* handler) {
    return CallWrapped(kIID_CompositionCompleted, handler,
                       [&](IUnknown* h) { return CreateCompositionOpt_Original(env, parent, opts, h); });
}

std::atomic<bool> g_envMethodsHooked{false};

// The environment's methods live in the WebView2 runtime and are shared by every
// environment, so hooking them once covers every window Outlook opens.
void HookEnvironment(IUnknown* env) {
    if (g_envMethodsHooked.exchange(true)) return;
    Wh_SetFunctionHook(Slot(env, kEnv_CreateController), (void*)CreateController_Hook, (void**)&CreateController_Original);
    IUnknown* env3 = nullptr;
    if (SUCCEEDED(env->QueryInterface(kIID_Env3, (void**)&env3))) {
        Wh_SetFunctionHook(Slot(env3, kEnv3_CreateCompositionController), (void*)CreateComposition_Hook,
                           (void**)&CreateComposition_Original);
        env3->Release();
    }
    IUnknown* env10 = nullptr;
    if (SUCCEEDED(env->QueryInterface(kIID_Env10, (void**)&env10))) {
        Wh_SetFunctionHook(Slot(env10, kEnv10_CreateControllerWithOptions), (void*)CreateControllerOpt_Hook,
                           (void**)&CreateControllerOpt_Original);
        Wh_SetFunctionHook(Slot(env10, kEnv10_CreateCompositionControllerWithOptions), (void*)CreateCompositionOpt_Hook,
                           (void**)&CreateCompositionOpt_Original);
        env10->Release();
    }
    Wh_ApplyHookOperations();
    Wh_Log(L"Environment methods hooked");
}

HRESULT OnEnvironment(HRESULT hr, IUnknown* env, IUnknown* outlook) {
    if (SUCCEEDED(hr) && env) HookEnvironment(env);
    return Forward(outlook, hr, env);
}

using CreateEnv_t = HRESULT(WINAPI*)(PCWSTR, PCWSTR, IUnknown*, IUnknown*);
CreateEnv_t CreateEnv_Original;

HRESULT WINAPI CreateEnv_Hook(PCWSTR browserFolder, PCWSTR userDataFolder, IUnknown* options, IUnknown* handler) {
    if (!handler) return CreateEnv_Original(browserFolder, userDataFolder, options, handler);
    auto* wrapped = new WrappedHandler(kIID_EnvCompleted, handler, OnEnvironment);
    HRESULT hr = CreateEnv_Original(browserFolder, userDataFolder, options, wrapped);
    wrapped->Release();
    return hr;
}

std::atomic<bool> g_loaderHooked{false};

bool HookLoader(bool apply) {
    HMODULE loader = GetModuleHandleW(L"WebView2Loader.dll");
    if (!loader || g_loaderHooked.exchange(true)) return false;
    void* target = (void*)GetProcAddress(loader, "CreateCoreWebView2EnvironmentWithOptions");
    if (!target) {
        Wh_Log(L"CreateCoreWebView2EnvironmentWithOptions not exported");
        return false;
    }
    Wh_SetFunctionHook(target, (void*)CreateEnv_Hook, (void**)&CreateEnv_Original);
    if (apply) Wh_ApplyHookOperations();
    Wh_Log(L"WebView2Loader hooked");
    return true;
}

using LoadLibraryExW_t = HMODULE(WINAPI*)(LPCWSTR, HANDLE, DWORD);
LoadLibraryExW_t LoadLibraryExW_Original;

// WebView2Loader.dll can arrive as a dependency of another DLL, so check after every load.
HMODULE WINAPI LoadLibraryExW_Hook(LPCWSTR name, HANDLE file, DWORD flags) {
    HMODULE module = LoadLibraryExW_Original(name, file, flags);
    if (module && !g_loaderHooked) HookLoader(true);
    return module;
}

BOOL Wh_ModInit() {
    BuildScript();
    if (HookLoader(false)) return TRUE;
    HMODULE kernelBase = GetModuleHandleW(L"kernelbase.dll");
    void* target = kernelBase ? (void*)GetProcAddress(kernelBase, "LoadLibraryExW") : nullptr;
    if (!target) return FALSE;
    Wh_SetFunctionHook(target, (void*)LoadLibraryExW_Hook, (void**)&LoadLibraryExW_Original);
    return TRUE;
}

void Wh_ModSettingsChanged() {
    BuildScript();
}

// Runs after Windhawk has removed the hooks, so no new handlers appear. Waits up to
// 5 seconds for WebView2 to release the ones it still holds before the DLL unloads.
void Wh_ModUninit() {
    for (int i = 0; i < 100 && g_liveObjects > 0; ++i) Sleep(50);
    if (g_liveObjects > 0) Wh_Log(L"Unloading while WebView2 still holds %d handlers", g_liveObjects.load());
}
