# go-excel-mcp-server# Go Excel MCP Server

A powerful MCP (Mark3labs Control Protocol) server that provides comprehensive Excel file manipulation capabilities using the excelize Go library.

## Features

### File Operations
- **Create new Excel workbooks**
- **Read data from worksheets**
- **Write data to worksheets**
- **Get detailed workbook metadata**

### Worksheet Management
- **Create new worksheets**
- **Delete existing worksheets**
- **Rename worksheets**

### Advanced Formatting
- **Comprehensive cell formatting** including:
  - Font styles (bold, italic, underline)
  - Text alignment and rotation
  - Number formatting (currency, percentages, dates)
  - Cell borders and colors
  - Background patterns and fills
  - Cell merging
  - Cell protection/locking
- **Conditional formatting support**

## Installation

1. Ensure you have Go installed (version 1.16 or higher recommended)
2. Clone the repository (if applicable)
3. Install dependencies:
   ```bash
   go mod tidy
   ```
4. Build the server:
   ```bash
   go build
   ```

## Usage

### Starting the Server
Run the compiled binary to start the MCP server in stdio mode:
```bash
./excel-tools-server
```

### MCP Tools Available

#### 1. Create Workbook
Creates a new Excel workbook at the specified path.

**Parameters:**
- `filepath` (string, required): Path where to create the new Excel file

**Example:**
```json
{"filepath": "output.xlsx"}
```

#### 2. Write Data to Excel
Writes tabular data to a specified worksheet starting from a given cell.

**Parameters:**
- `filepath` (string, required): Path to the Excel file
- `sheet_name` (string, required): Worksheet name
- `data` (array, required): List of lists (sublists are rows)
- `start_cell` (string, optional): Starting cell (default: "A1")

**Example:**
```json
{
  "filepath": "output.xlsx",
  "sheet_name": "Data",
  "data": [["Name", "Age"], ["Alice", 25], ["Bob", 30]],
  "start_cell": "B2"
}
```

#### 3. Read Data from Excel
Reads all data from a specified worksheet.

**Parameters:**
- `filepath` (string, required): Path to the Excel file
- `sheet_name` (string, required): Worksheet name to read

**Example:**
```json
{"filepath": "output.xlsx", "sheet_name": "Data"}
```

#### 4. Create Worksheet
Adds a new worksheet to an existing workbook.

**Parameters:**
- `filepath` (string, required): Path to the Excel file
- `sheet_name` (string, required): Name for the new worksheet

**Example:**
```json
{"filepath": "output.xlsx", "sheet_name": "NewSheet"}
```

#### 5. Delete Worksheet
Removes a specified worksheet from a workbook.

**Parameters:**
- `filepath` (string, required): Path to the Excel file
- `sheet_name` (string, required): Worksheet name to delete

**Example:**
```json
{"filepath": "output.xlsx", "sheet_name": "OldSheet"}
```

#### 6. Rename Worksheet
Changes the name of an existing worksheet.

**Parameters:**
- `filepath` (string, required): Path to the Excel file
- `old_name` (string, required): Current worksheet name
- `new_name` (string, required): New worksheet name

**Example:**
```json
{
  "filepath": "output.xlsx",
  "old_name": "Sheet1",
  "new_name": "MainData"
}
```

#### 7. Get Workbook Metadata
Retrieves metadata about a workbook including sheet list and ranges.

**Parameters:**
- `filepath` (string, required): Path to the Excel file
- `include_ranges` (boolean, optional): Whether to include range information

**Example:**
```json
{"filepath": "output.xlsx", "include_ranges": true}
```

#### 8. Format Range
Applies comprehensive formatting to a cell range.

**Parameters:**
- `filepath` (string, required): Path to the Excel file
- `sheet_name` (string, required): Worksheet name
- `start_cell` (string, required): Top-left cell of range
- `end_cell` (string, optional): Bottom-right cell (defaults to start_cell)
- Comprehensive formatting options including:
  - Font styles (bold, italic, underline)
  - Text alignment (horizontal/vertical)
  - Number formats
  - Cell borders and colors
  - Background patterns
  - Cell merging
  - Text wrapping
  - Cell protection

**Example:**
```json
{
  "filepath": "output.xlsx",
  "sheet_name": "Data",
  "start_cell": "A1",
  "end_cell": "D10",
  "bold": true,
  "font_size": 12,
  "bg_color": "FFFF00",
  "border_type": "thin",
  "number_format": "$#,##0.00"
}
```

## Error Handling

All tools return meaningful error messages in case of failures, including:
- File not found or inaccessible
- Invalid parameters
- Worksheet operations failures
- Formatting errors

## Dependencies

- [github.com/mark3labs/mcp-go](https://github.com/mark3labs/mcp-go) - MCP protocol implementation
- [github.com/xuri/excelize/v2](https://github.com/xuri/excelize) - Excel file manipulation library

## License

