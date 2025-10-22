Copy necessary fonts to Linux containers
----------------------------------------

XlsIO's Excel-to-PDF conversion on Linux uses the system fonts available inside the container. By default, only a few fonts are installed. Copy the required fonts to "/usr/local/share/fonts/" and refresh the font cache so they are used during conversion.

How to run this example
-----------------------

*   Download this project to your local disk.
*   Create a Fonts folder inside the project (or next to the solution).
*   Copy the required font files (.ttf/.otf) into the Fonts folder.
*   Open the solution in Visual Studio.
*   Rebuild the solution to restore/install NuGet packages.
*   Run the application.