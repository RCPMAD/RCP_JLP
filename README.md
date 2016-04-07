Jump List Parser (JLP)

Jump Lists are lists of recently opened items, such as files, folders, or websites, organized by yourself or the program that you use to open them. They have been around since Windows 7 and while many similar tools exist most do not support Windows 8, 8.1 and 10.

Originally this tool started as a fork of JumpLister which makes heavy use of a GUI and has a larger memory footprint. Now JLP looks nothing like it as the code has been rewritten, can run from the command line and does not come with a GUI. It is also aimed at incident responders that need to triage systems quickly. JPL is written in C# and uses a set of open source libraries (OpenMCDF 2.0, Shellify and CsvHelper) that can be found on GitHub.

Requirements
- Windows 8, 8.1 or 10
- .NET Framework Version 4

Features

- Can be executed from the command line interface. It is portable so it can be executed from any external flash drive. It does not come with a GUI therefore it has a low memory footprint.
- Detects and copies each user's jump list to the same location as the executable. This is done by creating a temp/username folder with the relevant username.
- Parses automatic and custom jump list destinations. These will be saved in a CSV file in the same location as the executable. New data will be appended if the CSV is not empty.
