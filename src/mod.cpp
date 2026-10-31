// ==WindhawkMod==
// @id              outlook-refined
// @name            Outlook Refined
// @description     Restyles new Outlook: favorites groups, per-account compose buttons, a resizable folder pane, an agenda with a month grid, the Other inbox under Focused, a tidier calendar and reading pane, and a print layout
// @version         0.5
// @author          Analytic Endeavors
// @github          https://github.com/ae-reid-havens
// @include         olk.exe
// @architecture    x86-64
// @license         MIT
// @compilerOptions -lole32 -loleaut32 -lurlmon -lbcrypt
// ==/WindhawkMod==

// ==WindhawkModReadme==
/*
# Outlook Refined

Restyles the new Outlook for Windows (olk.exe) from inside its own page. The mod
adds one script to each Outlook window when the window is created, and the
script stops at once unless the page is one of Outlook's own. It reads nothing
from disk, opens no ports and fetches nothing from the network.

![Folder pane](https://raw.githubusercontent.com/ae-reid-havens/outlook-refined/main/docs/folder-pane.png)

Restart Outlook after enabling the mod or changing a setting.

**What it asks Outlook for.** The agenda, month grid and Other inbox card read
your calendar and inbox through Outlook's own mail service, from inside its
page, signed with the session Outlook already holds: the mod keeps the request
headers Outlook sets, in memory, and reuses them. Those features refresh every
few minutes while they are on, and each is off until you turn it on. Nothing is
stored, logged or sent anywhere else, and no credential is written down.

**Where your edits are kept.** Groups, renames, folds, account order and the
pane and card sizes are saved in Windhawk's own storage for this mod, so they
leave when the mod does. State an earlier version wrote into Outlook's own local
storage is moved across the first time this one runs. Other Outlook windows
that are already open pick up a change when they are reopened.

**Folder pane:** Favorites in groups that fold (right-click a group header for
its menu, Shift+right-click a favorite to rename or move it), per-account
compose buttons, a resizable pane, account trees under one foldable heading,
accounts reordered by dragging.

**Agenda:** today and the next days from every account's calendar at the bottom
of the folder pane, with a join button for online meetings, under a month grid
shaded by booked hours. **Other inbox:** the newest Other conversations below
the Focused list.

![Agenda and the Other inbox card](https://raw.githubusercontent.com/ae-reid-havens/outlook-refined/main/docs/other-inbox.png)

**Calendar:** one row per account, compact time zones, start and end times on
event tiles, overlapping events cascaded or merged.

**Reading pane:** To and Cc on one line, a compact blocked-content bar, a
readable body width. **Ribbon:** the group-name row under the buttons goes.

**Print:** printing a message prints the message, not the whole window: the
panes, ribbon and buttons go, panes that scroll are released so nothing is cut
off, quoted replies can be lightened or dropped, pictures and tables are kept
whole, and anything too wide for the page is scaled down to fit.

**Known limits.** Most features find Outlook's own elements by English labels
("Focused", "Other", "Navigation pane", "Reading Pane") and read times written
as 1:30 PM, so in another display language, or with a 24-hour clock, those
parts quietly do nothing. The folder pane, the agenda and the print layout are
the parts that depend on this most.
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
- agenda: false
  $name: Agenda under the folder pane
  $description: Today and the next days from every account's calendar, with a join button for online meetings. Reads your calendar through Outlook's own mail service, signed with the session Outlook already holds, and refreshes every five minutes while it is on. Nothing is stored or sent anywhere else
- agendaDays: 3
  $name: Agenda days
  $description: Days the agenda lists, today included (1 to 14)
- monthHeatMap: false
  $name: Month grid above the agenda
  $description: Each day shaded by its booked hours; the arrows browse other months. Needs the agenda above to be on, and reads the same calendar data through Outlook's session
- otherCard: false
  $name: Other inbox under Focused
  $description: While an inbox shows its Focused tab, its newest Other conversations sit below the list. A button by the tabs switches back to Outlook's own tabs. Reads the Other list through Outlook's session and refreshes every two minutes while it is on
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
- printCompactHeader: true
  $name: Print, compact header
  $description: Drops the avatar and action buttons and tightens the sender block, so more of the message fits the first page
- printFit: overflow
  $name: Print, fit to the page width
  $options:
  - none: Never
  - overflow: Only when the message is too wide
  - always: Always
- printKeepChartsWhole: true
  $name: Print, keep pictures and tables whole
  $description: Never split a picture, table or code block across two pages
- printQuotedReplies: lighten
  $name: Print, quoted replies
  $options:
  - show: Print them as they are
  - lighten: Print them smaller and grey
  - hide: Leave them out
*/
// ==/WindhawkModSettings==

#include <unknwn.h>
#include <bcrypt.h>
#include <urlmon.h>

#include <atomic>
#include <map>
#include <string>
#include <string_view>
#include <vector>

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
constexpr GUID kIID_WebMessageReceived = {0x57213f19, 0x00e6, 0x49fa, {0x8e, 0x07, 0x89, 0x8e, 0xa0, 0x1e, 0xcb, 0xd2}};

enum {
    kEnv_CreateController = 3,
    kEnv3_CreateCompositionController = 9,
    kEnv10_CreateControllerWithOptions = 21,
    kEnv10_CreateCompositionControllerWithOptions = 22,
    kController_GetParentWindow = 21,
    kController_Close = 24,
    kController_GetCoreWebView2 = 25,
    kWebView_AddScriptOnDocumentCreated = 27,
    kWebView_RemoveScriptOnDocumentCreated = 28,
    kWebView_ExecuteScript = 29,
    kWebView_PostWebMessageAsString = 33,
    kWebView_AddWebMessageReceived = 34,
    kWebView_RemoveWebMessageReceived = 35,
    kArgs_GetSource = 3,
    kArgs_TryGetWebMessageAsString = 5,
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

// ==== the mod's own saved state ====

// Edits made inside Outlook (groups, renames, folds, account order, pane and card
// sizes) are the mod's state, not Outlook's, so they live in Windhawk's storage and
// leave with the mod. They are handed to the page in OLKT_CONFIG.store, and the page
// sends each change back over the WebView's message channel (see MessageReceived).
constexpr PCWSTR kStoreKeys[] = {
    L"olkt-favorites", L"olkt-account-order", L"olkt-folds",        L"olkt-pane-width",
    L"olkt-agenda-height", L"olkt-other-height", L"olkt-other-split", L"olkt-other-folds",
};
constexpr size_t kStoreMaxChars = 65536;

SRWLOCK g_storeLock = SRWLOCK_INIT;
std::map<std::wstring, std::wstring> g_store;

void LoadStore() {
    std::vector<wchar_t> buffer(kStoreMaxChars);
    AcquireSRWLockExclusive(&g_storeLock);
    g_store.clear();
    for (PCWSTR key : kStoreKeys) {
        buffer[0] = L'\0';
        if (Wh_GetStringValue(key, buffer.data(), buffer.size()) && buffer[0]) g_store[key] = buffer.data();
    }
    ReleaseSRWLockExclusive(&g_storeLock);
}

// The saved state as a JSON object, for the answer to a page that asks for it.
std::wstring StateJson() {
    std::wstring object = L"{";
    AcquireSRWLockShared(&g_storeLock);
    for (const auto& [key, value] : g_store) {
        if (object.size() > 1) object += L",";
        object += JsonString(key.c_str()) + L":" + JsonString(value.c_str());
    }
    ReleaseSRWLockShared(&g_storeLock);
    return object + L"}";
}

// A key made at load time and put in the page script, which no other page can read.
// Every message must carry it, so a page in another WebView, an add-in dialog for
// instance, cannot ask for the state or change it even though it shares the channel.
std::wstring g_key;

void MakeKey() {
    unsigned char bytes[16] = {};
    if (!BCRYPT_SUCCESS(BCryptGenRandom(nullptr, bytes, sizeof(bytes), BCRYPT_USE_SYSTEM_PREFERRED_RNG))) {
        // No randomness is a reason to carry no state, not a reason to accept anyone's.
        g_key.clear();
        return;
    }
    static constexpr wchar_t kHex[] = L"0123456789abcdef";
    g_key.clear();
    for (unsigned char b : bytes) {
        g_key += kHex[b >> 4];
        g_key += kHex[b & 0xF];
    }
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
    AddBool(json, L"printCompactHeader");
    AddString(json, L"printFit");
    AddBool(json, L"printKeepChartsWhole");
    AddString(json, L"printQuotedReplies");
    AddField(json, L"key", JsonString(g_key.c_str()));
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

// Set while the mod is being taken out: no new window is touched from here on.
std::atomic<bool> g_unloading{false};
// Set once the teardown has run. If anything could not be reached, the module is pinned
// in memory rather than unmapped, and whatever WebView2 still holds reads this and
// returns without touching the engine.
std::atomic<bool> g_unloaded{false};

// EventRegistrationToken from eventtoken.h: one INT64, so an add/remove pair can be
// called through the vtable without pulling in the WinRT headers.
struct EventToken {
    INT64 value;
};

// ==== state the page sends back ====

// The hosts Outlook itself serves, the same list the page script checks. The host comes
// from the URL parser, so it carries no user name, port or brackets, and an IDN host is
// already punycode; only a trailing dot and case are left to deal with.
bool OutlookHost(std::wstring_view host) {
    static constexpr std::wstring_view kHosts[] = {
        L"outlook.office.com",     L"outlook.office365.com",     L"outlook.live.com",
        L"outlook-sdf.office.com", L"outlook.cloud.microsoft",   L"outlook-sdf.cloud.microsoft",
    };
    if (!host.empty() && host.back() == L'.') host.remove_suffix(1);
    std::wstring lowered(host);
    for (wchar_t& c : lowered) c = (c >= L'A' && c <= L'Z') ? (wchar_t)(c - L'A' + L'a') : c;
    for (std::wstring_view known : kHosts) {
        if (lowered == known) return true;
        // A label boundary, so outlook.office.com.example.net is not a match.
        if (lowered.size() > known.size() + 1 && lowered.ends_with(known) &&
            lowered[lowered.size() - known.size() - 1] == L'.') {
            return true;
        }
    }
    return false;
}

// The document that sent the message, as the browser recorded it, not as the page says.
// Anything that is not https on one of Outlook's own hosts is refused, about:blank,
// data: and srcdoc documents included.
bool FromOutlook(PCWSTR url) {
    if (!url || !*url) return false;
    IUri* uri = nullptr;
    if (FAILED(CreateUri(url, Uri_CREATE_CANONICALIZE, 0, &uri)) || !uri) return false;
    BSTR scheme = nullptr;
    BSTR host = nullptr;
    // The scheme by name rather than by enum value: one less number to get wrong.
    bool https = SUCCEEDED(uri->GetSchemeName(&scheme)) && scheme &&
                 CompareStringOrdinal(scheme, -1, L"https", -1, TRUE) == CSTR_EQUAL;
    bool ok = https && SUCCEEDED(uri->GetHost(&host)) && host && OutlookHost(host);
    if (scheme) SysFreeString(scheme);
    if (host) SysFreeString(host);
    uri->Release();
    return ok;
}

// Writes the saved state. The caller has already checked who sent the message.
void StoreValue(std::wstring_view name, std::wstring_view text) {
    std::wstring key(name), value(text);
    // Only the mod's own keys, and nothing big enough to bloat the registry.
    bool known = false;
    for (PCWSTR k : kStoreKeys) known = known || key == k;
    if (!known || value.size() >= kStoreMaxChars) return;
    AcquireSRWLockExclusive(&g_storeLock);
    if (value.empty()) {
        g_store.erase(key);
        Wh_DeleteValue(key.c_str());
    } else {
        g_store[key] = value;
        Wh_SetStringValue(key.c_str(), value.c_str());
    }
    ReleaseSRWLockExclusive(&g_storeLock);
}

// A line the page sent, once the sender has been vouched for:
//   olkt-get <key>                 asks for the saved state, answered on the same channel
//   olkt-set <key> <name> <value>  writes one value; an empty value deletes it
// Every line carries the key from the script, so a page that never saw the script is
// turned away even on an Outlook host. Outlook's own messages are JSON and never match.
void HandleMessage(IUnknown* sender, PCWSTR text) {
    std::wstring_view message(text);
    bool get = message.starts_with(L"olkt-get\n");
    bool set = message.starts_with(L"olkt-set\n");
    if (!get && !set) return;
    message.remove_prefix(get ? 9 : 9);
    size_t line = message.find(L'\n');
    std::wstring_view key = line == std::wstring_view::npos ? message : message.substr(0, line);
    if (g_key.empty() || key != g_key) return;
    if (set) {
        if (line == std::wstring_view::npos) return;
        message.remove_prefix(line + 1);
        size_t name = message.find(L'\n');
        if (name == std::wstring_view::npos) return;
        StoreValue(message.substr(0, name), message.substr(name + 1));
        return;
    }
    std::wstring answer = L"olkt-state\n" + StateJson();
    using Post_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*, LPCWSTR);
    reinterpret_cast<Post_t>(Slot(sender, kWebView_PostWebMessageAsString))(sender, answer.c_str());
}

class MessageReceived final : public IUnknown {
    std::atomic<ULONG> refs_{1};

   public:
    MessageReceived() { ++g_liveObjects; }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** out) override {
        if (IsEqualGUID(iid, kIID_IUnknown) || IsEqualGUID(iid, kIID_WebMessageReceived)) {
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
            delete this;
            --g_liveObjects;
        }
        return n;
    }
    virtual HRESULT STDMETHODCALLTYPE Invoke(IUnknown* sender, IUnknown* args) {
        // The mod is on its way out, or gone and only still mapped: touch nothing.
        if (g_unloaded) return S_OK;
        using GetString_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*, LPWSTR*);
        LPWSTR source = nullptr;
        if (FAILED(reinterpret_cast<GetString_t>(Slot(args, kArgs_GetSource))(args, &source)) || !source) {
            return S_OK;
        }
        bool trusted = FromOutlook(source);
        CoTaskMemFree(source);
        if (!trusted) return S_OK;
        LPWSTR text = nullptr;
        // Not a string message (Outlook posts objects): nothing of ours.
        if (FAILED(reinterpret_cast<GetString_t>(Slot(args, kArgs_TryGetWebMessageAsString))(args, &text)) || !text) {
            return S_OK;
        }
        HandleMessage(sender, text);
        CoTaskMemFree(text);
        return S_OK;
    }
};

// ==== the WebViews our script is in ====

// One per live Outlook window, so disabling the mod can take the script out again
// instead of leaving it running until Outlook restarts. `controller` is kept as an
// identity only, never called, so it holds no reference and no window alive.
struct PageView {
    IUnknown* webview;
    void* controller;
    HWND window;
    std::wstring scriptId;
    EventToken token;
};

SRWLOCK g_viewsLock = SRWLOCK_INIT;
std::vector<PageView> g_views;

void ReleaseView(const PageView& view, bool removeScript);

void AddView(IUnknown* webview, void* controller, HWND window, PCWSTR scriptId) {
    // The teardown has already run for the windows it knew about, and this one arrived
    // too late to be on that list. Undo it here instead, on the thread it belongs to,
    // rather than leave a half-modded window behind.
    if (g_unloading) {
        PageView late{webview, controller, window, scriptId ? scriptId : L"", EventToken{}};
        webview->AddRef();  // ReleaseView drops it again
        ReleaseView(late, true);
        return;
    }
    MessageReceived* handler = new MessageReceived();
    EventToken token{};
    using Add_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*, IUnknown*, EventToken*);
    HRESULT hr = reinterpret_cast<Add_t>(Slot(webview, kWebView_AddWebMessageReceived))(webview, handler, &token);
    handler->Release();
    if (FAILED(hr)) Wh_Log(L"add_WebMessageReceived failed: 0x%08X", hr);
    webview->AddRef();
    AcquireSRWLockExclusive(&g_viewsLock);
    // Checked again under the lock: the teardown may have started while the handler was
    // being registered, in which case this view would never be torn down.
    if (g_unloading) {
        ReleaseSRWLockExclusive(&g_viewsLock);
        ReleaseView({webview, controller, window, scriptId ? scriptId : L"", token}, true);
        return;
    }
    g_views.push_back({webview, controller, window, scriptId ? scriptId : L"", token});
    ReleaseSRWLockExclusive(&g_viewsLock);
}

// Lets go of one WebView on its own thread: the script comes off, what is already on
// the page is undone, and the message handler goes with it.
void ReleaseView(const PageView& view, bool removeScript) {
    using RemoveScript_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*, LPCWSTR);
    using Execute_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*, LPCWSTR, IUnknown*);
    using RemoveMessage_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*, EventToken);
    if (removeScript) {
        if (!view.scriptId.empty()) {
            reinterpret_cast<RemoveScript_t>(Slot(view.webview, kWebView_RemoveScriptOnDocumentCreated))(
                view.webview, view.scriptId.c_str());
        }
        reinterpret_cast<Execute_t>(Slot(view.webview, kWebView_ExecuteScript))(
            view.webview, L"window.__olkt && window.__olkt.remove(true)", nullptr);
    }
    if (view.token.value) {
        reinterpret_cast<RemoveMessage_t>(Slot(view.webview, kWebView_RemoveWebMessageReceived))(view.webview,
                                                                                                 view.token);
    }
    view.webview->Release();
}

// Completion handler for AddScriptToExecuteOnDocumentCreated. Holds back Outlook's
// controller handler until the script is registered, because Outlook navigates from
// that handler and the window's first document must already get the script.
class ScriptAdded final : public IUnknown {
    std::atomic<ULONG> refs_{1};
    std::atomic<bool> forwarded_{false};
    IUnknown* outlook_;
    IUnknown* controller_;
    IUnknown* webview_;
    void* controllerKey_;
    HWND window_;

   public:
    ScriptAdded(IUnknown* outlook, IUnknown* controller, IUnknown* webview, void* controllerKey, HWND window)
        : outlook_(outlook), controller_(controller), webview_(webview), controllerKey_(controllerKey), window_(window) {
        outlook_->AddRef();
        controller_->AddRef();
        webview_->AddRef();
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
            webview_->Release();
            delete this;
            --g_liveObjects;
        }
        return n;
    }
    virtual HRESULT STDMETHODCALLTYPE Invoke(HRESULT hr, LPCWSTR id) {
        if (FAILED(hr)) {
            Wh_Log(L"Page script not added: hr=0x%08X", hr);
        } else {
            AddView(webview_, controllerKey_, window_, id);
        }
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

using ControllerClose_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*);
ControllerClose_t ControllerClose_Original;

// Outlook closes a window's controller when the window goes. Dropping our reference
// here keeps the mod from holding dead WebViews alive for the rest of the session.
HRESULT STDMETHODCALLTYPE ControllerClose_Hook(IUnknown* controller) {
    PageView view{};
    bool found = false;
    AcquireSRWLockExclusive(&g_viewsLock);
    for (auto it = g_views.begin(); it != g_views.end(); ++it) {
        if (it->controller == controller) {
            view = *it;
            found = true;
            g_views.erase(it);
            break;
        }
    }
    ReleaseSRWLockExclusive(&g_viewsLock);
    // Close is called on the WebView's own thread, so this is the right place for it.
    if (found) ReleaseView(view, false);
    return ControllerClose_Original(controller);
}

std::atomic<bool> g_controllerHooked{false};

// Controllers share one vtable, so this hook covers every window Outlook opens.
void HookControllerClose(IUnknown* controller) {
    if (g_controllerHooked.exchange(true)) return;
    Wh_SetFunctionHook(Slot(controller, kController_Close), (void*)ControllerClose_Hook,
                       (void**)&ControllerClose_Original);
    Wh_ApplyHookOperations();
    Wh_Log(L"Controller Close hooked");
}

// Adds the page script to each new WebView, then hands Outlook its controller once
// the script is registered (ScriptAdded). Any failure hands it over at once.
HRESULT OnController(HRESULT result_hr, IUnknown* result, IUnknown* outlook) {
    if (FAILED(result_hr) || !result) return Forward(outlook, result_hr, result);
    // Being taken out: hand Outlook its window untouched. This is the only place that
    // keeps a window from being half modded, because by the time the script completion
    // comes back the script has already run on the window's first document.
    if (g_unloading) return Forward(outlook, result_hr, result);
    IUnknown* controller = nullptr;
    if (FAILED(result->QueryInterface(kIID_Controller, (void**)&controller))) {
        Wh_Log(L"Controller interface missing");
        return Forward(outlook, result_hr, result);
    }
    IUnknown* webview = nullptr;
    using Get_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*, IUnknown**);
    HRESULT hr = reinterpret_cast<Get_t>(Slot(controller, kController_GetCoreWebView2))(controller, &webview);
    // The window the WebView lives in: taking the script out again later has to happen
    // on that window's own thread.
    HWND window = nullptr;
    using GetWindow_t = HRESULT(STDMETHODCALLTYPE*)(IUnknown*, HWND*);
    reinterpret_cast<GetWindow_t>(Slot(controller, kController_GetParentWindow))(controller, &window);
    HookControllerClose(controller);
    void* controllerKey = controller;
    controller->Release();
    if (FAILED(hr) || !webview) {
        Wh_Log(L"get_CoreWebView2 failed: 0x%08X", hr);
        return Forward(outlook, result_hr, result);
    }
    AcquireSRWLockShared(&g_scriptLock);
    std::wstring script = g_script;
    ReleaseSRWLockShared(&g_scriptLock);
    auto* done = new ScriptAdded(outlook, result, webview, controllerKey, window);
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

// ==== unload ====

// WebView2 objects answer only on the thread that made them, so the teardown is sent
// to each window's own thread, the way Windhawk's taskbar mods do it.
using RunFromWindowThreadProc_t = void(WINAPI*)(PVOID parameter);

struct RunFromWindowThreadParam {
    RunFromWindowThreadProc_t proc;
    PVOID param;
};

UINT g_runFromWindowThreadMessage;

LRESULT CALLBACK CallWndProcForRunFromWindowThread(int code, WPARAM wParam, LPARAM lParam) {
    if (code == HC_ACTION) {
        const CWPSTRUCT* cwp = (const CWPSTRUCT*)lParam;
        if (cwp->message == g_runFromWindowThreadMessage) {
            auto* param = (RunFromWindowThreadParam*)cwp->lParam;
            param->proc(param->param);
        }
    }
    return CallNextHookEx(nullptr, code, wParam, lParam);
}

// The parcel outlives the call on purpose. A timed-out SendMessageTimeout can still be
// delivered afterwards, and the hook procedure would then read a stack frame that has
// returned, so on that path the parcel is leaked rather than freed.
bool RunFromWindowThread(HWND window, RunFromWindowThreadProc_t proc, PVOID procParam) {
    DWORD thread = GetWindowThreadProcessId(window, nullptr);
    if (!thread) return false;
    if (thread == GetCurrentThreadId()) {
        proc(procParam);
        return true;
    }
    HHOOK hook = SetWindowsHookEx(WH_CALLWNDPROC, CallWndProcForRunFromWindowThread, nullptr, thread);
    if (!hook) return false;
    auto* parcel = new RunFromWindowThreadParam{proc, procParam};
    DWORD_PTR unused = 0;
    // A WebView2 UI thread can sit in a wait that pumps nothing, so the send is bounded.
    // Correctness does not rest on it: whatever is left over is made safe below.
    LRESULT sent = SendMessageTimeout(window, g_runFromWindowThreadMessage, 0, (LPARAM)parcel, SMTO_ABORTIFHUNG, 5000,
                                      &unused);
    UnhookWindowsHookEx(hook);
    if (!sent) return false;  // parcel deliberately leaked: a late delivery must still find it
    delete parcel;
    return true;
}

// One window's teardown, with its own copy of the view, for the same reason.
struct Teardown {
    PageView view;
    std::atomic<bool> done{false};
};

void WINAPI ReleaseViewOnThread(PVOID param) {
    auto* teardown = (Teardown*)param;
    if (!teardown->done.exchange(true)) ReleaseView(teardown->view, true);
}

// Keeps the mod's code mapped when something of ours is still reachable. Windhawk frees
// the DLL five seconds after the teardown whether or not WebView2 has let go, and a call
// into unmapped code takes Outlook down with it. One leaked module is the cheaper end.
void PinModule() {
    HMODULE self = nullptr;
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_PIN,
                       (LPCWSTR)&PinModule, &self);
}

// Runs while the hooks are still in place, so every open Outlook window gives the script
// back before the DLL goes: the styling, the timers and the wrapped fetch all stop at
// once, with no restart needed.
void Wh_ModBeforeUninit() {
    g_unloading = true;
    AcquireSRWLockExclusive(&g_viewsLock);
    std::vector<PageView> views;
    views.swap(g_views);
    ReleaseSRWLockExclusive(&g_viewsLock);  // released before any send, or the UI thread deadlocks
    bool leftBehind = false;
    for (const PageView& view : views) {
        auto* teardown = new Teardown{view};
        if (RunFromWindowThread(view.window, ReleaseViewOnThread, teardown)) {
            delete teardown;
            continue;
        }
        // Its window is gone or not answering. Leave the objects alone rather than call
        // into WebView2 from the wrong thread, and keep the code mapped for them.
        leftBehind = true;
        Wh_Log(L"Could not reach a WebView's thread; its script stays until Outlook restarts");
    }
    if (leftBehind) PinModule();
    g_unloaded = true;
}

BOOL Wh_ModInit() {
    g_runFromWindowThreadMessage = RegisterWindowMessage(L"Windhawk_RunFromWindowThread_outlook-refined");
    MakeKey();
    LoadStore();
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
    if (g_liveObjects > 0) {
        // WebView2 can still call into these, and the wait is over. Keeping the pages
        // mapped costs one module; unmapping them under a live caller costs Outlook.
        Wh_Log(L"Unloading while WebView2 still holds %d handlers; pinning the module",
               g_liveObjects.load());
        PinModule();
    }
}
