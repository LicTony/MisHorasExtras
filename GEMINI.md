# GEMINI Project Context: MisHorasExtras

## Project Overview

This is a .NET 8 WPF desktop application designed to interact with and manipulate a Microsoft Excel file named `MisHorasExtras.xlsm`. The application serves as a user interface to perform operations on this Excel file.

The core logic for Excel manipulation is encapsulated in the `ExcelManager.cs` class, which uses the `ClosedXML` library to provide a high-level API for reading, writing, and modifying `.xlsx` / `.xlsm` files.

**Key Technologies:**
*   **.NET 8:** The underlying framework.
*   **WPF:** Used for the graphical user interface.
*   **C#:** The primary programming language.
*   **ClosedXML:** A .NET library for creating and manipulating Excel files.

**Architecture:**
*   **`MainWindow.xaml` / `MainWindow.xaml.cs`:** The main window of the application, containing the UI and event handlers. The primary action is triggered by the `BtnEjecutar_Click` event.
*   **`ExcelManager.cs`:** A dedicated class that handles all interactions with the Excel file. It provides methods for opening, saving, reading from, and writing to the spreadsheet. This class is designed to be reusable and follows the `IDisposable` pattern for proper resource management.
*   **`MisHorasExtras.xlsm`:** An Excel file with macros that is included in the project and copied to the output directory. This file is the main target of the application's operations.

## Building and Running

To build and run this project, you will need Visual Studio or the .NET SDK.

**Using the .NET CLI:**

1.  **Restore Dependencies:**
    ```bash
    dotnet restore
    ```

2.  **Build the Project:**
    ```bash
    dotnet build
    ```

3.  **Run the Application:**
    ```bash
    dotnet run --project MisHorasExtras\MisHorasExtras.csproj
    ```

The executable will be generated in the `bin\Debug\net8.0-windows` directory, along with the `MisHorasExtras.xlsm` file.

## Development Conventions

*   **Excel Interaction:** All Excel-related operations should be performed through the `ExcelManager` class. Avoid using `ClosedXML` directly in the UI layer (`MainWindow.xaml.cs`).
*   **Resource Management:** The `ExcelManager` class implements `IDisposable`. Ensure that instances of `ExcelManager` are properly disposed of, for example, by using a `using` statement.
*   **Error Handling:** The `ExcelManager` class throws exceptions for invalid operations (e.g., file not found, empty tab name). These exceptions should be handled appropriately in the UI layer.
*   **User Interface:** The UI is defined in `MainWindow.xaml`. Changes to the UI should be made there.
*   **Event Handling:** UI event logic is located in the code-behind file `MainWindow.xaml.cs`.
