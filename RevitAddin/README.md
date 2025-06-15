# MyRevitAddin

This sample shows how to create a Revit 2022 add-in that adds a custom ribbon tab.

## Building

1. Open a Visual Studio project targeting .NET Framework 4.8.
2. Add references to `RevitAPI.dll` and `RevitAPIUI.dll` from your Revit 2022 installation.
3. Include `MyRevitAddin.cs` in the project and build a class library named `MyRevitAddin.dll`.
4. Copy the resulting DLL to `%AppData%\Autodesk\REVIT\Addins\2022\`.
5. Place `MyRevitAddin.addin` in the same directory.

Launching Revit will load the add-in and create a tab named **Custom Tab** with a dropdown and additional buttons.
