#define UNICODE
#define _UNICODE
#include <windows.h>
#include <string>
#include <vector>
#include <sstream>
#include <regex>

#define SCI_GETLENGTH            2006
#define SCI_GETTEXT              2182
#define SCI_SETTEXT              2181
#define SCI_BEGINUNDOACTION      2078
#define SCI_ENDUNDOACTION        2079
#define SCI_GETCURRENTPOS        2008
#define SCI_GETANCHOR            2009
#define SCI_GETFIRSTVISIBLELINE  2152
#define SCI_GETXOFFSET           2398
#define SCI_GOTOPOS              2025
#define SCI_SETSEL               2160
#define SCI_LINESCROLL           2300
#define SCI_SETXOFFSET           2397
#define SCI_LINEFROMPOSITION     2167
#define SCI_GETCOLUMN            2129
#define SCI_FINDCOLUMN           2456
#define SCI_POSITIONFROMLINE     2167  // alias, use SCI_LINEFROMPOSITION carefully
#undef  SCI_POSITIONFROMLINE
#define SCI_POSITIONFROMLINE     2194
#define NPPMSG                   (WM_USER + 1000)
#define NPPM_GETCURRENTSCINTILLA (NPPMSG + 4)

struct NppData {
    HWND _nppHandle;
    HWND _scintillaMainHandle;
    HWND _scintillaSecondHandle;
};

typedef void (*NppPluginFunc)();

#pragma pack(push, 8)
struct ShortcutKey {
    bool  _isAlt;
    bool  _isCtrl;
    bool  _isShift;
    UCHAR _key;
};
struct FuncItem {
    wchar_t       _itemName[64];
    NppPluginFunc _pFunc;
    int           _cmdID;
    bool          _init2Check;
    ShortcutKey*  _pShKey;
};
#pragma pack(pop)

HINSTANCE g_hModule = nullptr;
HWND      g_nppHwnd = nullptr;
HWND      g_sci1    = nullptr;
HWND      g_sci2    = nullptr;

struct Rule { std::wstring find, replace; bool isRegex; };
static std::vector<Rule> g_rules;

static std::wstring GetRulesPath() {
    wchar_t buf[MAX_PATH] = {};
    GetModuleFileNameW(g_hModule, buf, MAX_PATH);
    std::wstring p(buf);
    auto pos = p.rfind(L'\\');
    if (pos != std::wstring::npos) p = p.substr(0, pos + 1);
    return p + L"XReplace_rules.txt";
}

static std::wstring Utf8ToWide(const std::string& s) {
    if (s.empty()) return L"";
    int n = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, nullptr, 0);
    std::wstring w(n, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, &w[0], n);
    if (!w.empty() && w.back() == L'\0') w.pop_back();
    return w;
}

static std::string WideToUtf8(const std::wstring& w) {
    if (w.empty()) return "";
    int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string s(n, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, &s[0], n, nullptr, nullptr);
    if (!s.empty() && s.back() == '\0') s.pop_back();
    return s;
}

static std::wstring ParseEscapes(const std::wstring& s) {
    std::wstring out;
    for (size_t i = 0; i < s.size(); i++) {
        if (s[i] == L'\\' && i + 1 < s.size()) {
            switch (s[i+1]) {
                case L'n':  out += L'\n'; i++; break;
                case L't':  out += L'\t'; i++; break;
                case L'\\': out += L'\\'; i++; break;
                default:    out += s[i]; break;
            }
        } else { out += s[i]; }
    }
    return out;
}

static void LoadRules() {
    g_rules.clear();
    HANDLE hf = CreateFileW(GetRulesPath().c_str(), GENERIC_READ, FILE_SHARE_READ,
                            nullptr, OPEN_EXISTING, 0, nullptr);
    if (hf == INVALID_HANDLE_VALUE) return;
    DWORD sz = GetFileSize(hf, nullptr);
    std::string raw(sz, '\0');
    DWORD rd = 0;
    ReadFile(hf, &raw[0], sz, &rd, nullptr);
    CloseHandle(hf);
    if (raw.size() >= 3 &&
        (unsigned char)raw[0] == 0xEF &&
        (unsigned char)raw[1] == 0xBB &&
        (unsigned char)raw[2] == 0xBF)
        raw = raw.substr(3);
    std::wistringstream ss(Utf8ToWide(raw));
    std::wstring line;
    while (std::getline(ss, line)) {
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        if (line.empty() || line[0] == L'#') continue;
        std::vector<std::wstring> parts;
        size_t start = 0, pos = 0;
        while ((pos = line.find(L'|', start)) != std::wstring::npos) {
            parts.push_back(line.substr(start, pos - start));
            start = pos + 1;
        }
        parts.push_back(line.substr(start));
        if (parts.size() < 2) continue;
        Rule r;
        r.find    = ParseEscapes(parts[0]);
        r.replace = ParseEscapes(parts[1]);
        r.isRegex = (parts.size() >= 3 && parts[2] == L"regex");
        if (!r.find.empty()) g_rules.push_back(r);
    }
}

static std::wstring DoReplace(std::wstring text) {
    for (auto& rule : g_rules) {
        if (rule.isRegex) {
            try { std::wregex re(rule.find); text = std::regex_replace(text, re, rule.replace); }
            catch (...) {}
        } else {
            size_t pos = 0;
            while ((pos = text.find(rule.find, pos)) != std::wstring::npos) {
                text.replace(pos, rule.find.length(), rule.replace);
                pos += rule.replace.length();
            }
        }
    }
    return text;
}

static HWND GetCurrentScintilla() {
    int which = 0;
    SendMessage(g_nppHwnd, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&which);
    return (which == 0) ? g_sci1 : g_sci2;
}

void CmdReplaceAll();
void CmdOpenRules();

static ShortcutKey g_sk1 = {false, true, false, 0x31};  // Ctrl+1
static FuncItem g_funcItems[2];

static void InitFuncItems() {
    wcscpy_s(g_funcItems[0]._itemName, 64, L"Replace All");
    g_funcItems[0]._pFunc      = CmdReplaceAll;
    g_funcItems[0]._cmdID      = 0;
    g_funcItems[0]._init2Check = false;
    g_funcItems[0]._pShKey     = &g_sk1;

    wcscpy_s(g_funcItems[1]._itemName, 64, L"Edit Rules...");
    g_funcItems[1]._pFunc      = CmdOpenRules;
    g_funcItems[1]._cmdID      = 0;
    g_funcItems[1]._init2Check = false;
    g_funcItems[1]._pShKey     = nullptr;
}

void CmdReplaceAll() {
    LoadRules();
    if (g_rules.empty()) {
        MessageBoxW(g_nppHwnd,
            (L"No rules found.\nPlease check:\n" + GetRulesPath()).c_str(),
            L"XReplace", MB_OK | MB_ICONWARNING);
        return;
    }
    HWND sci = GetCurrentScintilla();
    if (!sci) return;
    int len = (int)SendMessage(sci, SCI_GETLENGTH, 0, 0);
    if (len <= 0) return;
    std::string buf(len + 1, '\0');
    SendMessage(sci, SCI_GETTEXT, (WPARAM)(len + 1), (LPARAM)buf.data());
    buf.resize(len);
    std::wstring wtext    = Utf8ToWide(buf);
    std::wstring replaced = DoReplace(wtext);
    if (replaced == wtext) return;
    std::string u8 = WideToUtf8(replaced);

    // 保存光标和滚动位置（用行列号，避免全角→半角字节数变化导致偏移错误）
    int curBytePos    = (int)SendMessage(sci, SCI_GETCURRENTPOS,       0, 0);
    int anchorBytePos = (int)SendMessage(sci, SCI_GETANCHOR,           0, 0);
    int firstLine     = (int)SendMessage(sci, SCI_GETFIRSTVISIBLELINE, 0, 0);
    int xOffset       = (int)SendMessage(sci, SCI_GETXOFFSET,         0, 0);

    // 把字节偏移转换为宽字符列号，保存行号和列号
    // 先把当前 UTF-8 文本转成宽字符，然后从字节偏移反推宽字符列
    auto BytePosToLineCol = [&](int bytePos, int& outLine, int& outCol) {
        // 直接用 Scintilla 的行列消息（它内部维护 UTF-8，行列以字符为单位）
        outLine = (int)SendMessage(sci, SCI_LINEFROMPOSITION, (WPARAM)bytePos, 0);
        outCol  = (int)SendMessage(sci, SCI_GETCOLUMN,        (WPARAM)bytePos, 0);
    };
    auto LineColToBytePos = [&](int line, int col) -> int {
        return (int)SendMessage(sci, SCI_FINDCOLUMN, (WPARAM)line, (LPARAM)col);
    };

    int curLine, curCol, ancLine, ancCol;
    BytePosToLineCol(curBytePos,    curLine, curCol);
    BytePosToLineCol(anchorBytePos, ancLine, ancCol);

    SendMessage(sci, SCI_BEGINUNDOACTION, 0, 0);
    SendMessage(sci, SCI_SETTEXT, 0, (LPARAM)u8.c_str());
    SendMessage(sci, SCI_ENDUNDOACTION, 0, 0);

    // 先把光标移到文档开头，防止 Scintilla 自动滚动到旧光标位置
    SendMessage(sci, SCI_GOTOPOS, 0, 0);

    // 恢复滚动位置
    SendMessage(sci, SCI_LINESCROLL, 0, firstLine - (int)SendMessage(sci, SCI_GETFIRSTVISIBLELINE, 0, 0));
    SendMessage(sci, SCI_SETXOFFSET, xOffset, 0);

    // 用行列号恢复光标（行列号不受字节数变化影响）
    int newLen    = (int)SendMessage(sci, SCI_GETLENGTH, 0, 0);
    int newCur    = LineColToBytePos(curLine, curCol);
    int newAnchor = LineColToBytePos(ancLine, ancCol);
    if (newCur    > newLen) newCur    = newLen;
    if (newAnchor > newLen) newAnchor = newLen;
    SendMessage(sci, SCI_SETSEL, (WPARAM)newAnchor, (LPARAM)newCur);
}

void CmdOpenRules() {
    std::wstring path = GetRulesPath();
    if (GetFileAttributesW(path.c_str()) == INVALID_FILE_ATTRIBUTES) {
        const char* def =
            "\xEF\xBB\xBF"
            "# XReplace rules\n"
            "# Format: find|replace\n"
            "# Format: find|replace|regex\n"
            "# Lines starting with # are comments\n"
            "#\n"
            "\xEF\xBC\x9F|?\n"
            "\xEF\xBC\x9B|,\n"
            "\xEF\xBC\x9A|:\n"
            "\xEF\xBC\x8C|,\n"
            "\xEF\xBC\x89|)\n"
            "\xEF\xBC\x88|(\n"
            "\xe3\x80\x82|,\n"
            "\xe3\x80\x81|/\n"
            "\xEF\xBC\x81|!\n"
            "\xe2\x80\x94\xe2\x80\x94|\xe2\x80\x94\n"
            "\xE2\x80\x9C|\"\n"
            "\xE2\x80\x9D|\"\n"
            "\xE2\x80\x98|\"\n"
            "\xE2\x80\x99|\"\n"
            "\\*\\*||regex\n";
        HANDLE hf = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, 0, nullptr);
        if (hf != INVALID_HANDLE_VALUE) {
            DWORD w; WriteFile(hf, def, (DWORD)strlen(def), &w, nullptr);
            CloseHandle(hf);
        }
    }
    ShellExecuteW(nullptr, L"open", path.c_str(), nullptr, nullptr, SW_SHOW);
}

extern "C" __declspec(dllexport)
void setInfo(NppData data) {
    g_nppHwnd = data._nppHandle;
    g_sci1    = data._scintillaMainHandle;
    g_sci2    = data._scintillaSecondHandle;
}

extern "C" __declspec(dllexport)
const wchar_t* getName() { return L"XReplace"; }

extern "C" __declspec(dllexport)
FuncItem* getFuncsArray(int* nbF) {
    InitFuncItems();
    *nbF = 2;
    return g_funcItems;
}

extern "C" __declspec(dllexport) void beNotified(void*) {}
extern "C" __declspec(dllexport) LRESULT messageProc(UINT, WPARAM, LPARAM) { return TRUE; }
extern "C" __declspec(dllexport) BOOL isUnicode() { return TRUE; }

BOOL APIENTRY DllMain(HMODULE hMod, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) { g_hModule = hMod; DisableThreadLibraryCalls(hMod); }
    return TRUE;
}
