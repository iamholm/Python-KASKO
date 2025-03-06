#include <windows.h>
#include <shellapi.h>
#include <string>

// укажите имя вашего Python EXE
static std::string pythonExe = "Casco15op.exe";  

static const char* CLASS_NAME = "MySplashClass";
static HWND g_hwnd = NULL;
static UINT_PTR g_timerId = 0;

// Процедура окна
LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CREATE:
    {
        // Запустим ваш Python-EXE
        ShellExecuteA(NULL, "open", pythonExe.c_str(), NULL, NULL, SW_SHOWDEFAULT);

        // Создаём текст (статик) внутри окна
        CreateWindowA(
            "STATIC",
            "Загрузка... Пожалуйста, подождите 5 секунд",
            WS_CHILD | WS_VISIBLE | SS_CENTER,
            10, 10, 280, 40,
            hWnd, NULL, ((LPCREATESTRUCT)lParam)->hInstance, NULL
        );

        // Таймер на 5 сек (5000 мс)
        g_timerId = SetTimer(hWnd, 1, 5000, NULL);
        return 0;
    }
    case WM_TIMER:
        if (wParam == 1)
        {
            KillTimer(hWnd, g_timerId);
            DestroyWindow(hWnd); // Закрываем окно
        }
        return 0;
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcA(hWnd, msg, wParam, lParam);
}

// Функция, создающая окно
int RunSplash(HINSTANCE hInstance)
{
    // 1) Регистрируем класс окна
    WNDCLASSA wc = {0};
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
    RegisterClassA(&wc);

    // 2) Создаём окно (с рамкой)
    HWND hWnd = CreateWindowA(
        CLASS_NAME,
        "Loading...", // заголовок окна
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
        CW_USEDEFAULT, CW_USEDEFAULT, 300, 120,
        NULL, NULL, hInstance, NULL
    );

    if (!hWnd)
    {
        MessageBoxA(NULL, "Failed to create window", "Error", MB_ICONERROR);
        return -1;
    }

    // Центрируем
    {
        RECT rc;
        GetWindowRect(hWnd, &rc);
        int winW = rc.right - rc.left;
        int winH = rc.bottom - rc.top;

        int scrW = GetSystemMetrics(SM_CXSCREEN);
        int scrH = GetSystemMetrics(SM_CYSCREEN);
        int posX = (scrW - winW) / 2;
        int posY = (scrH - winH) / 2;

        SetWindowPos(hWnd, NULL, posX, posY, winW, winH, 0);
    }

    ShowWindow(hWnd, SW_SHOW);
    UpdateWindow(hWnd);

    // Цикл сообщений
    MSG msg;
    while (GetMessageA(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessageA(&msg);
    }
    return (int)msg.wParam;
}

// Точка входа
int WINAPI WinMain(HINSTANCE hInst, HINSTANCE, LPSTR, int)
{
    return RunSplash(hInst);
}
