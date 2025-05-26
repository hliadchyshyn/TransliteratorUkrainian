# Excel to CSV Ukrainian Transliterator

This is a simple C# console application that reads data from an Excel `.xlsx` file, transliterates any Ukrainian Cyrillic text into Latin characters, and writes the output to a `.csv` file.

## Features

- Reads `.xlsx` files using [ClosedXML](https://github.com/ClosedXML/ClosedXML)
- Transliterates Ukrainian Cyrillic characters to Latin
- Outputs to a UTF-8 encoded `.csv` file
- Supports multi-row and multi-column Excel sheets

## Requirements

- .NET 6.0 or later
- [ClosedXML](https://www.nuget.org/packages/ClosedXML/)

## Installation

1. Clone this repository or copy the code.
2. Install the required NuGet package:


## Build
1. Place your input Excel file in the same directory as the executable and name it input.xlsx.

2. Run the program:
dotnet run
3. The output will be saved as output.csv in the same directory.


## Usage
Place your input Excel file in the same directory as the executable and name it input.xlsx.

## Example
Input (input.xlsx)
| Ім'я   | Прізвище |
| ------ | -------- |
| Андрій | Шевченко |


Output (output.csv)
Andrii,Shevchenko

## Transliteration Rules
This tool uses a custom character-by-character mapping based on official Ukrainian-to-Latin transliteration standards. Examples:

| Cyrillic | Latin |
| -------- | ----- |
| А        | A     |
| Б        | B     |
| Ж        | Zh    |
| Ш        | Sh    |
| Ю        | Yu    |
| Я        | Ya    |


## License

This project is open-source and free to use.
