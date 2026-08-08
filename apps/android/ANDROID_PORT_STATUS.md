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
- Android-specific mobile shell and back-button handling.
- Shared AI provider layer with OpenRouter/NVIDIA-compatible routing.
- Android platform adapter installed without Electron IPC.

### Major phase completed: Docs vertical slice
- The real shared GenOffice Docs renderer is now bundled into the Android application.
- The existing `@genoffice/docx-engine` parser/save pipeline is reused instead of creating a second document engine.
- DOCX open uses the Android/Web file picker and stores an editable cache copy through Capacitor Filesystem.
- DOCX manual/new saves write to the Android Documents area and expose the result through the Android share sheet.
- Crash-recovery copies use Capacitor cache storage.
- Image insertion uses the Android file picker.
- Shared AI provider settings and chat are available to the Docs renderer through an Android `DesktopApi` adapter.
- PDF fragment merging uses `pdf-lib` in the Android renderer.
- Touch/viewport overrides are isolated to Android so the desktop renderer is not modified for mobile.
- GitHub Actions builds the Android APK from this branch.

### Remaining major phases
1. Harden the Docs adapter: true Android document URI persistence, SAF integration, attachment import, direct PDF export/share, and multi-document tabs.
2. Sheets Android editor + XLSX processing adapter.
3. Slides Android editor + PPTX import/export adapter.
4. PDF Android viewer/annotation/export adapter.
5. Shared platform contract so Electron and Capacitor implement the same capabilities explicitly rather than treating `DesktopApi` as Android's long-term abstraction.
6. Phone/tablet responsive UX pass and hardware keyboard support.
7. Physical-device smoke tests and signed release AAB/APK.

A build that only displays the mobile shell must not be described as the finished GenOffice Android app. The Android branch now contains a functional Docs editor vertical slice; the remaining office apps are still separate completion gates.
