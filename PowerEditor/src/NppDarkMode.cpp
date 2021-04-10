#include "nppDarkMode.h"

#include "DarkMode/IatHook.h"
#include "DarkMode/UAHMenuBar.h"

#include <Shlwapi.h>

#include <Uxtheme.h>
#include <vssym32.h>

#include <unordered_set>
#include <mutex>

#pragma comment(lib, "uxtheme.lib")

enum IMMERSIVE_HC_CACHE_MODE
{
	IHCM_USE_CACHED_VALUE,
	IHCM_REFRESH
};

enum WINDOWCOMPOSITIONATTRIB
{
	WCA_UNDEFINED = 0,
	WCA_NCRENDERING_ENABLED = 1,
	WCA_NCRENDERING_POLICY = 2,
	WCA_TRANSITIONS_FORCEDISABLED = 3,
	WCA_ALLOW_NCPAINT = 4,
	WCA_CAPTION_BUTTON_BOUNDS = 5,
	WCA_NONCLIENT_RTL_LAYOUT = 6,
	WCA_FORCE_ICONIC_REPRESENTATION = 7,
	WCA_EXTENDED_FRAME_BOUNDS = 8,
	WCA_HAS_ICONIC_BITMAP = 9,
	WCA_THEME_ATTRIBUTES = 10,
	WCA_NCRENDERING_EXILED = 11,
	WCA_NCADORNMENTINFO = 12,
	WCA_EXCLUDED_FROM_LIVEPREVIEW = 13,
	WCA_VIDEO_OVERLAY_ACTIVE = 14,
	WCA_FORCE_ACTIVEWINDOW_APPEARANCE = 15,
	WCA_DISALLOW_PEEK = 16,
	WCA_CLOAK = 17,
	WCA_CLOAKED = 18,
	WCA_ACCENT_POLICY = 19,
	WCA_FREEZE_REPRESENTATION = 20,
	WCA_EVER_UNCLOAKED = 21,
	WCA_VISUAL_OWNER = 22,
	WCA_HOLOGRAPHIC = 23,
	WCA_EXCLUDED_FROM_DDA = 24,
	WCA_PASSIVEUPDATEMODE = 25,
	WCA_USEDARKMODECOLORS = 26,
	WCA_LAST = 27
};

struct WINDOWCOMPOSITIONATTRIBDATA
{
	WINDOWCOMPOSITIONATTRIB Attrib;
	PVOID pvData;
	SIZE_T cbData;
};

static FARPROC ResolveFunction(LPCTSTR dllname, const char* name, int ordinal = -1)
{
	const auto mod = GetModuleHandleW(dllname);
	if (mod)
	{
		const auto by_name = name ? GetProcAddress(mod, name) : nullptr;
		if (by_name)
			return by_name;
		const auto by_ord = ordinal != -1 ? GetProcAddress(mod, MAKEINTRESOURCEA(ordinal)) : nullptr;
		if (by_ord)
			return by_ord;
	}
	return nullptr;
}

// a wrapper template so we can use funny types in macros
template <typename T>
using wrap_t = T;

#define MAKE_FN_RESOLVER(dllname, name, ordinal, ...) wrap_t<__VA_ARGS__> FORCEINLINE fn ## name () \
	{\
		static auto ptr = ResolveFunction(TEXT(#dllname), # name, ordinal);\
		return (wrap_t<__VA_ARGS__>)ptr;\
	}

MAKE_FN_RESOLVER(ntdll, RtlGetNtVersionNumbers, -1, void(WINAPI*)(LPDWORD major, LPDWORD minor, LPDWORD build));
MAKE_FN_RESOLVER(user32, SetWindowCompositionAttribute, -1, BOOL(WINAPI*)(HWND hWnd, WINDOWCOMPOSITIONATTRIBDATA*));

// 1809 17763
MAKE_FN_RESOLVER(uxtheme.dll, ShouldAppsUseDarkMode, 132, bool(WINAPI*)())
MAKE_FN_RESOLVER(uxtheme.dll, AllowDarkModeForWindow, 133, bool (WINAPI*)(HWND hWnd, bool allow))
MAKE_FN_RESOLVER(uxtheme.dll, FlushMenuThemes, 136, void (WINAPI*)())
MAKE_FN_RESOLVER(uxtheme.dll, RefreshImmersiveColorPolicyState, 104, void (WINAPI*)())
MAKE_FN_RESOLVER(uxtheme.dll, IsDarkModeAllowedForWindow, 137, bool (WINAPI*)(HWND hWnd))
MAKE_FN_RESOLVER(uxtheme.dll, GetIsImmersiveColorUsingHighContrast, 106, bool (WINAPI*)(IMMERSIVE_HC_CACHE_MODE mode))
MAKE_FN_RESOLVER(uxtheme.dll, OpenNcThemeData, 49, HTHEME(WINAPI*)(HWND hWnd, LPCWSTR pszClassList))

// 1903 18362
MAKE_FN_RESOLVER(uxtheme.dll, ShouldSystemUseDarkMode, 138, bool (WINAPI*)())
MAKE_FN_RESOLVER(uxtheme.dll, IsDarkModeAllowedForApp, 139, bool (WINAPI*)())

bool g_darkModeSupported = false;
bool g_darkModeEnabled = false;
DWORD g_buildNumber = 0;

// wrapper to decide which version is available. In reality code generated *should* be the same for both branches
static void AllowDarkModeForApp(bool allow)
{
	// 1903 18362
	enum PreferredAppMode
	{
		Default,
		AllowDark,
		ForceDark,
		ForceLight,
		Max
	};

	using pfnSetPreferredAppMode = PreferredAppMode(WINAPI*)(PreferredAppMode AppMode);
	using pfnAllowDarkModeForApp = BOOL(WINAPI*)(BOOL Allow);

	static auto ord135 = ResolveFunction(L"uxtheme.dll", nullptr, 135);
	if (!ord135)
		return;

	if (g_buildNumber < 18362)
		reinterpret_cast<pfnAllowDarkModeForApp>(ord135)(allow);
	else
		reinterpret_cast<pfnSetPreferredAppMode>(ord135)(allow ? AllowDark : Default);
}

static bool AllowDarkModeForWindow(HWND hWnd, bool allow)
{
	if (g_darkModeSupported && fnAllowDarkModeForWindow())
		return fnAllowDarkModeForWindow()(hWnd, allow);
	return false;
}

static bool IsHighContrast()
{
	HIGHCONTRASTW highContrast = {sizeof(highContrast)};
	if (SystemParametersInfoW(SPI_GETHIGHCONTRAST, sizeof(highContrast), &highContrast, FALSE))
		return highContrast.dwFlags & HCF_HIGHCONTRASTON;
	return false;
}

void RefreshTitleBarThemeColor(HWND hWnd)
{
	BOOL dark = FALSE;
	if (fnIsDarkModeAllowedForWindow()(hWnd) &&
		fnShouldAppsUseDarkMode()() &&
		!IsHighContrast())
	{
		dark = TRUE;
	}
	if (g_buildNumber < 18362)
		SetPropW(hWnd, L"UseImmersiveDarkModeColors", reinterpret_cast<HANDLE>(static_cast<INT_PTR>(dark)));
	else if (fnSetWindowCompositionAttribute())
	{
		WINDOWCOMPOSITIONATTRIBDATA data = {WCA_USEDARKMODECOLORS, &dark, sizeof(dark)};
		fnSetWindowCompositionAttribute()(hWnd, &data);
	}
}

bool IsColorSchemeChangeMessage(LPARAM lParam)
{
	bool is = false;
	if (lParam && (0 == lstrcmpiW(reinterpret_cast<LPCWCH>(lParam), L"ImmersiveColorSet")))
	{
		fnRefreshImmersiveColorPolicyState()();
		is = true;
	}
	fnGetIsImmersiveColorUsingHighContrast()(IHCM_REFRESH);
	return is;
}

bool IsColorSchemeChangeMessage(UINT message, LPARAM lParam)
{
	if (message == WM_SETTINGCHANGE)
		return IsColorSchemeChangeMessage(lParam);
	return false;
}

// limit dark scroll bar to specific windows and their children

std::unordered_set<HWND> g_darkScrollBarWindows;
std::mutex g_darkScrollBarMutex;

void EnableDarkScrollBarForWindowAndChildren(HWND hwnd)
{
	std::lock_guard<std::mutex> lock(g_darkScrollBarMutex);
	g_darkScrollBarWindows.insert(hwnd);
}

bool IsWindowOrParentUsingDarkScrollBar(HWND hwnd)
{
	const HWND hwndRoot = GetAncestor(hwnd, GA_ROOT);

	std::lock_guard<std::mutex> lock(g_darkScrollBarMutex);
	if (g_darkScrollBarWindows.count(hwnd))
		return true;

	if (hwnd != hwndRoot && g_darkScrollBarWindows.count(hwndRoot))
		return true;

	return false;
}

static bool UnprotectingMemcpy(void* dst, const void* src, size_t size)
{
	DWORD oldProtect;
	if (!VirtualProtect(dst, size, PAGE_EXECUTE_READWRITE, &oldProtect))
		return false;
	memcpy(dst, src, size);
	VirtualProtect(dst, size, oldProtect, &oldProtect);
	return true; // if the last VirtualProtect failed we will just be left on rwx, don't care, patch is done.
}

void FixDarkScrollBar()
{
	const HMODULE hComctl = LoadLibraryExW(L"comctl32.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
	if (hComctl)
	{
		auto addr = FindDelayLoadThunkInModule(hComctl, "uxtheme.dll", 49); // OpenNcThemeData
		if (addr)
		{
			const auto MyOpenThemeData = [](HWND hWnd, LPCWSTR classList) -> HTHEME
			{
				if (wcscmp(classList, L"ScrollBar") == 0)
				{
					if (IsWindowOrParentUsingDarkScrollBar(hWnd))
					{
						hWnd = nullptr;
						classList = L"Explorer::ScrollBar";
					}
				}
				return fnOpenNcThemeData()(hWnd, classList);
			};

			const auto ptr = reinterpret_cast<ULONG_PTR>(static_cast<decltype(fnOpenNcThemeData())>(MyOpenThemeData));
			const auto dst = &addr->u1.Function;

			UnprotectingMemcpy(dst, &ptr, sizeof(ptr));
		}
	}
}

constexpr bool CheckBuildNumber(DWORD buildNumber)
{
	return buildNumber >= 17763; // 2004
}

void InitDarkMode()
{
	if (fnRtlGetNtVersionNumbers())
	{
		DWORD major = 0, minor = 0;
		fnRtlGetNtVersionNumbers()(&major, &minor, &g_buildNumber);
		g_buildNumber &= ~0xF0000000;
		if (major == 10 && minor == 0 && CheckBuildNumber(g_buildNumber))
		{
			const HMODULE hUxtheme = LoadLibraryExW(L"uxtheme.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
			if (hUxtheme)
			{
				if (fnOpenNcThemeData() &&
					fnRefreshImmersiveColorPolicyState() &&
					fnShouldAppsUseDarkMode() &&
					fnAllowDarkModeForWindow() &&
					fnIsDarkModeAllowedForWindow())
				{
					g_darkModeSupported = true;

					AllowDarkModeForApp(true);
					fnRefreshImmersiveColorPolicyState()();

					g_darkModeEnabled = fnShouldAppsUseDarkMode()() && !IsHighContrast();

					FixDarkScrollBar();
				}
			}
		}
	}
}

namespace NppDarkMode
{
	bool IsEnabled()
	{
		return g_darkModeEnabled;
	}

	COLORREF InvertLightness(COLORREF color)
	{
		WORD h = 0;
		WORD s = 0;
		WORD l = 0;
		ColorRGBToHLS(color, &h, &l, &s);

		l = 240 - l;

		const COLORREF invert_c = ColorHLSToRGB(h, l, s);

		return invert_c;
	}

	COLORREF InvertLightnessSofter(COLORREF color)
	{
		WORD h = 0;
		WORD s = 0;
		WORD l = 0;
		ColorRGBToHLS(color, &h, &l, &s);

		l = min(240 - l, 211);

		const COLORREF invert_c = ColorHLSToRGB(h, l, s);

		return invert_c;
	}

	COLORREF GetBackgroundColor()
	{
		return RGB(0x20, 0x20, 0x20);
	}

	COLORREF GetSofterBackgroundColor()
	{
		return RGB(0x30, 0x30, 0x30);
	}

	COLORREF GetTextColor()
	{
		return RGB(0xE0, 0xE0, 0xE0);
	}

	COLORREF GetDarkerTextColor()
	{
		return RGB(0xC0, 0xC0, 0xC0);
	}

	COLORREF GetEdgeColor()
	{
		return RGB(0x80, 0x80, 0x80);
	}

	HBRUSH GetBackgroundBrush()
	{
		static HBRUSH g_hbrBackground = ::CreateSolidBrush(GetBackgroundColor());
		return g_hbrBackground;
	}

	HBRUSH GetSofterBackgroundBrush()
	{
		static HBRUSH g_hbrSofterBackground = ::CreateSolidBrush(GetSofterBackgroundColor());
		return g_hbrSofterBackground;
	}

	// handle events

	bool OnSettingChange(HWND hwnd, LPARAM lParam) // true if dark mode toggled
	{
		bool toggled = false;
		if (IsColorSchemeChangeMessage(lParam))
		{
			const bool darkModeWasEnabled = g_darkModeEnabled;
			g_darkModeEnabled = fnShouldAppsUseDarkMode()() && !IsHighContrast();

			NppDarkMode::RefreshTitleBarThemeColor(hwnd);

			if (!!darkModeWasEnabled != !!g_darkModeEnabled)
			{
				toggled = true;
				RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN);
			}
		}

		return toggled;
	}

	// processes messages related to UAH / custom menubar drawing.
	// return true if handled, false to continue with normal processing in your wndproc
	bool UAHWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, LRESULT* lr)
	{
		if (!NppDarkMode::IsEnabled())
			return false;

		static HTHEME g_menuTheme = nullptr;

		UNREFERENCED_PARAMETER(wParam);
		switch (message)
		{
		case WM_UAHDRAWMENU:
		{
			auto pUDM = (UAHMENU*)lParam;
			RECT rc = {0};

			// get the menubar rect
			{
				MENUBARINFO mbi = {sizeof(mbi)};
				GetMenuBarInfo(hWnd, OBJID_MENU, 0, &mbi);

				RECT rcWindow;
				GetWindowRect(hWnd, &rcWindow);

				// the rcBar is offset by the window rect
				rc = mbi.rcBar;
				OffsetRect(&rc, -rcWindow.left, -rcWindow.top);

				rc.top -= 1;
			}

			FillRect(pUDM->hdc, &rc, NppDarkMode::GetBackgroundBrush());

			*lr = 0;

			return true;
		}
		case WM_UAHDRAWMENUITEM:
		{
			auto pUDMI = (UAHDRAWMENUITEM*)lParam;

			// get the menu item string
			wchar_t menuString[256] = {0};
			MENUITEMINFO mii = {sizeof(mii), MIIM_STRING};
			{
				mii.dwTypeData = menuString;
				mii.cch = (sizeof(menuString) / 2) - 1;

				GetMenuItemInfo(pUDMI->um.hmenu, pUDMI->umi.iPosition, TRUE, &mii);
			}

			// get the item state for drawing

			DWORD dwFlags = DT_CENTER | DT_SINGLELINE | DT_VCENTER;

			int iTextStateID = MPI_NORMAL;
			int iBackgroundStateID = MPI_NORMAL;
			{
				if ((pUDMI->dis.itemState & ODS_INACTIVE) | (pUDMI->dis.itemState & ODS_DEFAULT))
				{
					// normal display
					iTextStateID = MPI_NORMAL;
					iBackgroundStateID = MPI_NORMAL;
				}
				if (pUDMI->dis.itemState & ODS_HOTLIGHT)
				{
					// hot tracking
					iTextStateID = MPI_HOT;
					iBackgroundStateID = MPI_HOT;
				}
				if (pUDMI->dis.itemState & ODS_SELECTED)
				{
					// clicked -- MENU_POPUPITEM has no state for this, though MENU_BARITEM does
					iTextStateID = MPI_HOT;
					iBackgroundStateID = MPI_HOT;
				}
				if ((pUDMI->dis.itemState & ODS_GRAYED) || (pUDMI->dis.itemState & ODS_DISABLED))
				{
					// disabled / grey text
					iTextStateID = MPI_DISABLED;
					iBackgroundStateID = MPI_DISABLED;
				}
				if (pUDMI->dis.itemState & ODS_NOACCEL)
				{
					dwFlags |= DT_HIDEPREFIX;
				}
			}

			if (!g_menuTheme)
			{
				g_menuTheme = OpenThemeData(hWnd, L"Menu");
			}

			if (iBackgroundStateID == MPI_NORMAL || iBackgroundStateID == MPI_DISABLED)
			{
				FillRect(pUDMI->um.hdc, &pUDMI->dis.rcItem, NppDarkMode::GetBackgroundBrush());
			}
			else
			{
				DrawThemeBackground(g_menuTheme, pUDMI->um.hdc, MENU_POPUPITEM, iBackgroundStateID, &pUDMI->dis.rcItem, nullptr);
			}
			DTTOPTS dttopts = {sizeof(dttopts)};
			if (iTextStateID == MPI_NORMAL || iTextStateID == MPI_HOT)
			{
				dttopts.dwFlags |= DTT_TEXTCOLOR;
				dttopts.crText = NppDarkMode::GetTextColor();
			}
			DrawThemeTextEx(g_menuTheme, pUDMI->um.hdc, MENU_POPUPITEM, iTextStateID, menuString, mii.cch, dwFlags, &pUDMI->dis.rcItem, &dttopts);

			*lr = 0;

			return true;
		}
		case WM_THEMECHANGED:
		{
			if (g_menuTheme)
			{
				CloseThemeData(g_menuTheme);
				g_menuTheme = nullptr;
			}
			// continue processing in main wndproc
			return false;
		}
		default:
			return false;
		}
	}

	// from DarkMode.h

	void InitDarkMode()
	{
		::InitDarkMode();
	}

	void AllowDarkModeForApp(bool allow)
	{
		::AllowDarkModeForApp(allow);
	}

	bool AllowDarkModeForWindow(HWND hWnd, bool allow)
	{
		return ::AllowDarkModeForWindow(hWnd, allow);
	}

	void RefreshTitleBarThemeColor(HWND hWnd)
	{
		::RefreshTitleBarThemeColor(hWnd);
	}

	void EnableDarkScrollBarForWindowAndChildren(HWND hwnd)
	{
		::EnableDarkScrollBarForWindowAndChildren(hwnd);
	}
}
