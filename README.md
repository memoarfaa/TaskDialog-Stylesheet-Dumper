# TaskDialog Stylesheet Dumper

A Windows GUI tool that extracts the TaskDialog stylesheet from `comctl32.dll` (resource ID 2455, type `UIFILE`), displays the XML, and evaluates all theme function calls (`gtf`, `gtc`, `gtmar`, `gtmet`, `dtb`) using the `uxtheme` API. The output shows the actual font, color, margin, and metric values from the current Windows visual style.

## Features
- Loads the correct themed version of `comctl32.dll` via a manifest (uses WinSxS).
- Parses the `UIFILE` resource (XML or binary) with `XmlLite`.
- Extracts the `<style resid="TaskDialog">` block and displays it.
- Lists every theme function call found in the stylesheet.
- Evaluates each call with `uxtheme` and shows the resolved values (font name, RGB color, margins, etc.).
- Saves the entire output (including evaluated values) as a UTF‑16 XML file.
- Simple resizable GUI with an edit control and a Save button.

## Requirements
- Windows Vista or later (tested on Windows 10/11).
- Visual Studio 2022 (or any compiler that supports Windows SDK, e.g., MinGW with appropriate libraries).
- Visual Styles must be enabled (the tool will still work but theme evaluation will show fallbacks).

## Building

### Using Visual Studio 2022
1. Open `TaskDialog-Stylesheet-Dumper.sln`.
2. Select the desired configuration (Release/Debug, x86/x64).
3. Build the solution.

