# ExcelCellInfoGenerator

## 📦Description

This project is a Maven project that writes cell information from an Excel file to a Json file.

## 🗒Sample of the output Json file

<details>
  <summary>Check the contents of the sample file.</summary>

```json
[
  {
    "SheetName": "シート1",
    "CellInfo": [
      {
        "cellRange": "A1",
        "value": "テスト",
        "alignment": "GENERAL",
        "borderBottom": "THIN",
        "borderLeft": "THIN",
        "borderRight": "THIN",
        "borderTop": "THIN",
        "bottomBorderColor": "0",
        "leftBorderColor": "0",
        "rightBorderColor": "0",
        "topBorderColor": "0",
        "dataFormat": "General",
        "fillBackgroundColor": null,
        "fillForegroundColor": null,
        "fillPattern": "NO_FILL",
        "font": {
          "bold": "false",
          "charSet": "0",
          "xssfColor": null,
          "themeColor": "0",
          "fontHeight": "220",
          "fontHeightInPoints": "11",
          "fontName": "Calibri",
          "italic": "false",
          "strikeout": "false",
          "typeOffset": "0",
          "underline": "0"
        },
        "hidden": "false",
        "indention": "0",
        "locked": "true",
        "rotation": "0",
        "verticalAlignment": "BOTTOM"
      },
      {
        "cellRange": "B1",
        "value": "サンプル",
        "alignment": "GENERAL",
        "borderBottom": "NONE",
        "borderLeft": "NONE",
        "borderRight": "NONE",
        "borderTop": "NONE",
        "bottomBorderColor": "8",
        "leftBorderColor": "8",
        "rightBorderColor": "8",
        "topBorderColor": "8",
        "dataFormat": "General",
        "fillBackgroundColor": "[-1, -1, 0]",
        "fillForegroundColor": "[-1, -1, 0]",
        "fillPattern": "SOLID_FOREGROUND",
        "font": {
          "bold": "false",
          "charSet": "0",
          "xssfColor": null,
          "themeColor": "0",
          "fontHeight": "220",
          "fontHeightInPoints": "11",
          "fontName": "Calibri",
          "italic": "false",
          "strikeout": "false",
          "typeOffset": "0",
          "underline": "0"
        },
        "hidden": "false",
        "indention": "0",
        "locked": "true",
        "rotation": "0",
        "verticalAlignment": "BOTTOM"
      },
      {
        "cellRange": "A2",
        "value": "Excel",
        "alignment": "GENERAL",
        "borderBottom": "NONE",
        "borderLeft": "NONE",
        "borderRight": "NONE",
        "borderTop": "NONE",
        "bottomBorderColor": "8",
        "leftBorderColor": "8",
        "rightBorderColor": "8",
        "topBorderColor": "8",
        "dataFormat": "General",
        "fillBackgroundColor": null,
        "fillForegroundColor": null,
        "fillPattern": "NO_FILL",
        "font": {
          "bold": "false",
          "charSet": "0",
          "xssfColor": "[0, 0, 0]",
          "themeColor": "1",
          "fontHeight": "220",
          "fontHeightInPoints": "11",
          "fontName": "Arial",
          "italic": "false",
          "strikeout": "false",
          "typeOffset": "0",
          "underline": "0"
        },
        "hidden": "false",
        "indention": "0",
        "locked": "true",
        "rotation": "0",
        "verticalAlignment": "BOTTOM"
      },
      {
        "cellRange": "B3",
        "value": "",
        "alignment": "GENERAL",
        "borderBottom": "NONE",
        "borderLeft": "NONE",
        "borderRight": "NONE",
        "borderTop": "NONE",
        "bottomBorderColor": "8",
        "leftBorderColor": "8",
        "rightBorderColor": "8",
        "topBorderColor": "8",
        "dataFormat": "m/d/yy",
        "fillBackgroundColor": null,
        "fillForegroundColor": null,
        "fillPattern": "NO_FILL",
        "font": {
          "bold": "false",
          "charSet": "0",
          "xssfColor": null,
          "themeColor": "0",
          "fontHeight": "220",
          "fontHeightInPoints": "11",
          "fontName": "Calibri",
          "italic": "false",
          "strikeout": "false",
          "typeOffset": "0",
          "underline": "0"
        },
        "hidden": "false",
        "indention": "0",
        "locked": "true",
        "rotation": "0",
        "verticalAlignment": "BOTTOM"
      }
    ]
  }
]
```

</details>

## 💬Usage

### Clone the Git project

```shell
git clone https://github.com/no-dumps/excel-cell-info-generator.git
```

### Maven install

#### Mac

```shell
./mvnw install
```

#### Windows

```shell
mvnw install
```

### Execution

#### Mac

```shell
./mvnw exec:java -DexcelPath=<Excel file path>
```

#### Windows

```shell
mvnw exec:java -DexcelPath=<Excel file path>
```

## 🎫 License

- [MIT](https://raw.githubusercontent.com/no-dumps/excel-cell-info-generator/main/LICENSE)

## 👀Author

- [GitHub](https://github.com/NaokiOouchi)
- [Twitter](https://twitter.com/NaoNoaNaoNoaN)