#include <Windows.h>
#include <vsstyle.h>
#include <Uxtheme.h>
#pragma comment(lib, "Uxtheme")

#define WM_UAHDRAWMENU         0x0091
#define WM_UAHDRAWMENUITEM     0x0092
#define WM_UAHMEASUREMENUITEM  0x0094

static HBRUSH g_MenuBarBg = CreateSolidBrush(RGB(0x0, 0x0, 0x0));
static HBRUSH g_MenuItemSelected = CreateSolidBrush(RGB(0x22, 0x22, 0x22));

//=================================================================================
// Credits: https://github.com/adzm/win32-custom-menubar-aero-theme

typedef union tagUAHMENUITEMMETRICS
{
    struct {
        DWORD cx;
        DWORD cy;
    } rgsizeBar[2];
    struct {
        DWORD cx;
        DWORD cy;
    } rgsizePopup[4];
} UAHMENUITEMMETRICS;

typedef struct tagUAHMENUPOPUPMETRICS
{
    DWORD rgcx[4];
    DWORD fUpdateMaxWidths : 2;
} UAHMENUPOPUPMETRICS;

typedef struct tagUAHMENU
{
    HMENU hmenu;
    HDC hdc;
    DWORD dwFlags;
} UAHMENU;

typedef struct tagUAHMENUITEM
{
    int iPosition;
    UAHMENUITEMMETRICS umim;
    UAHMENUPOPUPMETRICS umpm;
} UAHMENUITEM;

typedef struct UAHDRAWMENUITEM
{
    DRAWITEMSTRUCT dis;
    UAHMENU um;
    UAHMENUITEM umi;
} UAHDRAWMENUITEM;

//=================================================================================

static WNDPROC g_WindowProc = nullptr;

BOOL CALLBACK ChildWindowsEnum(HWND p_Window, LPARAM p_LPARAM)
{
    if ((GetWindowLongA(p_Window, GWL_STYLE) & BS_RADIOBUTTON) != 0) {
        SetWindowTheme(p_Window, L"", L"");
    }

    return TRUE;
}

RECT MapRectFromClientToWndCoords(HWND hwnd, const RECT& r)
{
    RECT wnd_coords = r;
    MapWindowPoints(hwnd, NULL, reinterpret_cast<POINT*>(&wnd_coords), 2);

    RECT scr_coords;
    GetWindowRect(hwnd, &scr_coords);
    OffsetRect(&wnd_coords, -scr_coords.left, -scr_coords.top);

    return wnd_coords;
}

RECT GetNonclientMenuBorderRect(HWND hwnd)
{
    RECT r;
    GetClientRect(hwnd, &r);
    r = MapRectFromClientToWndCoords(hwnd, r);
    int y = r.top - 1;
    return { r.left, y, r.right, y + 1 };
}

void MenuBarUnderlineDraw(HWND p_Window)
{
    HDC hdc = GetWindowDC(p_Window);
    RECT r = GetNonclientMenuBorderRect(p_Window);
    FillRect(hdc, &r, g_MenuBarBg);
    ReleaseDC(p_Window, hdc);
}

LRESULT WindowProc(HWND p_Window, UINT p_Msg, WPARAM p_WParam, LPARAM p_LParam)
{
    switch (p_Msg)
    {
        case WM_UAHDRAWMENU:
        {
            UAHMENU* pUDM = (UAHMENU*)p_LParam;
            RECT rc = { 0 };

            {
                MENUBARINFO mbi = { sizeof(mbi) };
                GetMenuBarInfo(p_Window, OBJID_MENU, 0, &mbi);

                RECT rcWindow;
                GetWindowRect(p_Window, &rcWindow);

                rc = mbi.rcBar;
                OffsetRect(&rc, -rcWindow.left, -rcWindow.top);
            }

            FillRect(pUDM->hdc, &rc, g_MenuBarBg);

            return true;
        }
        case WM_UAHDRAWMENUITEM:
        {
            UAHDRAWMENUITEM* pUDMI = (UAHDRAWMENUITEM*)p_LParam;

            HBRUSH* pbrBackground = &g_MenuBarBg;

            wchar_t menuString[256] = { 0 };
            MENUITEMINFO mii = { sizeof(mii), MIIM_STRING };
            {
                mii.dwTypeData = menuString;
                mii.cch = (sizeof(menuString) / 2) - 1;

                GetMenuItemInfo(pUDMI->um.hmenu, pUDMI->umi.iPosition, TRUE, &mii);
            }

            DWORD dwFlags = DT_CENTER | DT_SINGLELINE | DT_VCENTER;

            int iTextStateID = 0;
            int iBackgroundStateID = 0;
            {
                if ((pUDMI->dis.itemState & ODS_INACTIVE) | (pUDMI->dis.itemState & ODS_DEFAULT)) {
                    // normal display
                    iTextStateID = MPI_NORMAL;
                    iBackgroundStateID = MPI_NORMAL;
                }
                if (pUDMI->dis.itemState & ODS_HOTLIGHT) {
                    // hot tracking
                    iTextStateID = MPI_HOT;
                    iBackgroundStateID = MPI_HOT;

                    pbrBackground = &g_MenuItemSelected;
                }
                if (pUDMI->dis.itemState & ODS_SELECTED) {
                    iTextStateID = MPI_HOT;
                    iBackgroundStateID = MPI_HOT;

                    pbrBackground = &g_MenuItemSelected;
                }
                if ((pUDMI->dis.itemState & ODS_GRAYED) || (pUDMI->dis.itemState & ODS_DISABLED)) {
                    iTextStateID = MPI_DISABLED;
                    iBackgroundStateID = MPI_DISABLED;
                }
                if (pUDMI->dis.itemState & ODS_NOACCEL) {
                    dwFlags |= DT_HIDEPREFIX;
                }
            }

            static HTHEME g_menuTheme = nullptr;
            if (!g_menuTheme) {
                g_menuTheme = OpenThemeData(p_Window, L"Menu");
            }

            DTTOPTS opts = { sizeof(opts), DTT_TEXTCOLOR, iTextStateID != MPI_DISABLED ? RGB(0xFF, 0xFF, 0xFF) : RGB(0xFF, 0xFF, 0xFF) };

            FillRect(pUDMI->um.hdc, &pUDMI->dis.rcItem, *pbrBackground);
            DrawThemeTextEx(g_menuTheme, pUDMI->um.hdc, MENU_BARITEM, MBI_NORMAL, menuString, mii.cch, dwFlags, &pUDMI->dis.rcItem, &opts);

            return true;
        }
        case WM_PAINT:
        {
            EnumChildWindows(p_Window, ChildWindowsEnum, 0);
            MenuBarUnderlineDraw(p_Window);
        }
        break; 
        case WM_NCPAINT:
        {
            LRESULT _Res = DefWindowProcA(p_Window, WM_NCPAINT, p_WParam, p_LParam);
            MenuBarUnderlineDraw(p_Window);
            return _Res;
        }
        case WM_ACTIVATEAPP:
        {
            UpdateWindow(p_Window);
            InvalidateRect(p_Window, nullptr, false);
        }
        break;
    }

    return CallWindowProcA(g_WindowProc, p_Window, p_Msg, p_WParam, p_LParam);
}

//=================================================================================

DWORD __stdcall MainThread(void* p_Reserved)
{
    uintptr_t _CalcAddress = reinterpret_cast<uintptr_t>(GetModuleHandleA(0));
    HWND* _CalcFrame = reinterpret_cast<HWND*>(_CalcAddress + 0x785C0);

    while (!*_CalcFrame) {
        Sleep(1);
    }

    HWND _CalcWindow = GetParent(*_CalcFrame);

    g_WindowProc = reinterpret_cast<WNDPROC>(SetWindowLongPtrA(_CalcWindow, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(WindowProc)));

    HMENU _CalcMenu = 0;
    if (!_CalcMenu) {
        _CalcMenu = GetMenu(_CalcWindow);
    }

    return 0;
}

//=================================================================================

int __stdcall DllMain(HMODULE p_Module, DWORD p_Reason, void* p_Reserved)
{
    if (p_Reason == DLL_PROCESS_ATTACH)
    {
        DisableThreadLibraryCalls(p_Module);
        HANDLE _Thread = CreateThread(0, 0, MainThread, 0, 0, 0);
        if (_Thread) {
            CloseHandle(_Thread);
        }
    }

    return 1;
}
