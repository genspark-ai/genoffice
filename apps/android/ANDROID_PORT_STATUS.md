# GenOffice Android Port

## Target
A real Android app, not a website wrapper: Docs, Sheets, Slides and PDF with Android file handling, shared AI providers, touch support and no Electron runtime.

## Architecture

```text
Shared React/TypeScript editors + engine packages
                    |
          +---------+---------+
          |                   |
       Electron            Capacitor
          |                   |
   Windows/macOS          Android APK/AAB
```

## Current status

### Bootstrapped
- Capacitor Android application target.
- Android package id: `com.genoffice.mobile`.
- Android-specific web entry point.
- Android back-button handling.
- Keyboard/status-bar integration.
- Separate mobile shell so desktop code is not modified.

### Not yet complete
The existing editor renderers are not drop-in mobile components. They currently depend on `DesktopApi`/Electron IPC for file open/save, AI streaming, attachments, PDF export, search, lifecycle and native dialogs. The desktop shell also hosts editors as Electron `WebContentsView` children. These APIs must be replaced by a browser/Capacitor platform adapter before the Android editor can be considered functional.

### Completion gates
1. Platform adapter contract shared by Electron and Android.
2. Docs open/edit/save/recovery/import image/AI.
3. PDF open/annotate/export/share.
4. Slides touch editing/import/export/AI.
5. Sheets Android-safe XLSX processing replacing desktop sidecar process assumptions.
6. OpenRouter + NVIDIA NIM AI route on Android.
7. Android storage picker and share flows.
8. Phone/tablet responsive editor layouts.
9. Device smoke tests on physical Android hardware.
10. Signed release AAB/APK build.

A build that only displays the mobile shell must not be described as the finished GenOffice Android app.
