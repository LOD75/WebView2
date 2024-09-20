// compile with: /D_UNICODE /DUNICODE /DWIN32 /D_WINDOWS /c

#include <windows.h>
#include <stdlib.h>
#include <string>
#include <tchar.h>
#include <wrl.h>
#include <wil/com.h>
// <IncludeHeader>
// include WebView2 header
#include "WebView2.h"
// </IncludeHeader>

using namespace Microsoft::WRL;

// Global variables

// The main window class name.
static TCHAR szWindowClass[] = _T("DesktopApp");

// The string that appears in the application's title bar.
static TCHAR szTitle[] = _T("WebView sample");

HINSTANCE hInst;

// Forward declarations of functions included in this code module:
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

// Pointer to WebViewController
static wil::com_ptr<ICoreWebView2Controller> webviewController;

static wil::com_ptr<ICoreWebView2Environment6> environment;

// Pointer to WebView window
static wil::com_ptr<ICoreWebView2_16> webview;

int CALLBACK WinMain(
	_In_ HINSTANCE hInstance,
	_In_ HINSTANCE hPrevInstance,
	_In_ LPSTR     lpCmdLine,
	_In_ int       nCmdShow
)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);
	wcex.style = CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc = WndProc;
	wcex.cbClsExtra = 0;
	wcex.cbWndExtra = 0;
	wcex.hInstance = hInstance;
	wcex.hIcon = LoadIcon(hInstance, IDI_APPLICATION);
	wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
	wcex.lpszMenuName = NULL;
	wcex.lpszClassName = szWindowClass;
	wcex.hIconSm = LoadIcon(wcex.hInstance, IDI_APPLICATION);

	if (!RegisterClassEx(&wcex))
	{
		MessageBox(NULL,
			_T("Call to RegisterClassEx failed!"),
			_T("Windows Desktop Guided Tour"),
			NULL);

		return 1;
	}

	// Store instance handle in our global variable
	hInst = hInstance;

	// The parameters to CreateWindow explained:
	// szWindowClass: the name of the application
	// szTitle: the text that appears in the title bar
	// WS_OVERLAPPEDWINDOW: the type of window to create
	// CW_USEDEFAULT, CW_USEDEFAULT: initial position (x, y)
	// 500, 100: initial size (width, length)
	// NULL: the parent of this window
	// NULL: this application does not have a menu bar
	// hInstance: the first parameter from WinMain
	// NULL: not used in this application
	HWND hWnd = CreateWindow(
		szWindowClass,
		szTitle,
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, CW_USEDEFAULT,
		1200, 900,
		NULL,
		NULL,
		hInstance,
		NULL
	);

	if (!hWnd)
	{
		MessageBox(NULL,
			_T("Call to CreateWindow failed!"),
			_T("Windows Desktop Guided Tour"),
			NULL);

		return 1;
	}

	// The parameters to ShowWindow explained:
	// hWnd: the value returned from CreateWindow
	// nCmdShow: the fourth parameter from WinMain
	ShowWindow(hWnd,
		nCmdShow);
	UpdateWindow(hWnd);

	// <-- WebView2 sample code starts here -->
	// Step 3 - Create a single WebView within the parent window
	// Locate the browser and set up the environment for WebView
	CreateCoreWebView2EnvironmentWithOptions(nullptr, nullptr, nullptr,
		Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[hWnd](HRESULT result, ICoreWebView2Environment* env) -> HRESULT {

				// Create a CoreWebView2Controller and get the associated CoreWebView2 whose parent is the main window hWnd
				env->CreateCoreWebView2Controller(hWnd, Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
					[hWnd](HRESULT result, ICoreWebView2Controller* controller) -> HRESULT {
						if (controller != nullptr) {
							webviewController = controller;
							wil::com_ptr<ICoreWebView2> webview2;
							webviewController->get_CoreWebView2(&webview2);
							webview2.query_to(&webview);
						}

						// Add a few settings for the webview
						// The demo step is redundant since the values are the default settings
						wil::com_ptr<ICoreWebView2Settings> settings;
						webview->get_Settings(&settings);
						settings->put_IsScriptEnabled(TRUE);
						settings->put_AreDefaultScriptDialogsEnabled(TRUE);
						settings->put_IsWebMessageEnabled(TRUE);


						/*wil::com_ptr<ICoreWebView2PrintSettings> printSettings;
						environment->CreatePrintSettings(&printSettings);

						wil::com_ptr<ICoreWebView2PrintSettings2> printSettings2;
						printSettings->QueryInterface(IID_PPV_ARGS(&printSettings2));

						printSettings->put_Orientation(COREWEBVIEW2_PRINT_ORIENTATION_PORTRAIT);*/

						// Resize WebView to fit the bounds of the parent window
						RECT bounds;
						GetClientRect(hWnd, &bounds);
						webviewController->put_Bounds(bounds);

						//Navigate to html document
						webview->Navigate(L"C:\\WebView2Sample\\WebView2Samples\\GettingStartedGuides\\Win32_GettingStarted\\HTMLPage1.html");

						//ULONG size = (file.GetLength() + 1) * (sizeof(WCHAR));
						const wchar_t* characters = L"C:\\WebView2Sample\\WebView2Samples\\GettingStartedGuides\\Win32_GettingStarted\\HTMLPage1.html";
						ULONG size = wcslen(characters) + 1;
						LPWSTR temp = new wchar_t[size];
						wcsncpy_s(temp,size, characters, size -1);
						wil::unique_cotaskmem_string title = wil::make_cotaskmem_string(temp, size);
						webview->get_DocumentTitle(&title);

						webview->Print(nullptr,
							//printSettings.get(),
							Callback<ICoreWebView2PrintCompletedHandler>(
								[title = std::move(title)](HRESULT errorCode, COREWEBVIEW2_PRINT_STATUS printStatus)->HRESULT
						{
							std::wstring message = L"";
							if (errorCode == S_OK && printStatus == COREWEBVIEW2_PRINT_STATUS_SUCCEEDED)
							{
								message = L"Printing " + std::wstring(title.get()) +
									L" document to printer is succeeded";
							}
							else if (
								errorCode == S_OK &&
								printStatus == COREWEBVIEW2_PRINT_STATUS_PRINTER_UNAVAILABLE)
							{
								message = L"Selected printer is not found, not available, offline or "
									L"error state.";
							}
							else if (errorCode == E_INVALIDARG)
							{
								message = L"Invalid settings provided for the specified printer";
							}
							else if (errorCode == E_ABORT)
							{
								message = L"Printing " + std::wstring(title.get()) +
									L" document already in progress";
							}
							else
							{
								message = L"Printing " + std::wstring(title.get()) +
									L" document to printer is failed";
							}

							//AsyncMessageBox(message, L"Print to printer");

							return S_OK;
						}).Get());
						return S_OK;
					}).Get());
				return S_OK;
			}).Get());


	
	// <-- WebView2 sample code ends here -->

	// Main message loop:
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	return (int)msg.wParam;
}

//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_DESTROY  - post a quit message and return
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	TCHAR greeting[] = _T("Hello, Windows desktop!");

	switch (message)
	{
	case WM_SIZE:
		if (webviewController != nullptr) {
			RECT bounds;
			GetClientRect(hWnd, &bounds);
			webviewController->put_Bounds(bounds);
		};
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
		break;
	}

	return 0;
}
