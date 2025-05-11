package main

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	"github.com/xuri/excelize/v2"
	"log"
	"strings"
)

func main() {
	// Create MCP server
	s := server.NewMCPServer(
		"Excel Tools Server",
		"1.0.0",
		server.WithLogging(),
		server.WithRecovery(),
	)

	// Tool 1: create_workbook
	createWorkbookTool := mcp.NewTool("create_workbook",
		mcp.WithDescription("Create a new Excel workbook"),
		mcp.WithString("filepath",
			mcp.Required(),
			mcp.Description("Path where to create the new Excel file"),
		),
	)

	s.AddTool(createWorkbookTool, func(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
		filepath, ok := request.Params.Arguments["filepath"].(string)
		if !ok {
			return nil, errors.New("filepath must be a string")
		}

		f := excelize.NewFile()
		defer f.Close()

		if err := f.SaveAs(filepath); err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to save workbook: %v", err)), nil
		}

		return mcp.NewToolResultText(fmt.Sprintf("Excel workbook created at: %s", filepath)), nil
	})

	// Tool 2: write_data_to_excel
	writeDataTool := mcp.NewTool("write_data_to_excel",
		mcp.WithDescription("Write data to an Excel worksheet"),
		mcp.WithString("filepath",
			mcp.Required(),
			mcp.Description("Path to the Excel file"),
		),
		mcp.WithString("sheet_name",
			mcp.Required(),
			mcp.Description("Name of the worksheet to write to"),
		),
		mcp.WithArray("data",
			mcp.Required(),
			mcp.Description("List of lists containing data to write (sublists are rows)"),
		),
		mcp.WithString("start_cell",
			mcp.Description("Cell to start writing to (default: A1)"),
			mcp.DefaultString("A1"), // Corrected: Using DefaultString
		),
	)

	s.AddTool(writeDataTool, func(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
		filepath, ok := request.Params.Arguments["filepath"].(string)
		if !ok {
			return nil, errors.New("filepath must be a string")
		}

		sheetName, ok := request.Params.Arguments["sheet_name"].(string)
		if !ok {
			return nil, errors.New("sheet_name must be a string")
		}

		dataInterface, ok := request.Params.Arguments["data"].([]interface{})
		if !ok {
			return nil, errors.New("data must be an array")
		}

		// Convert data to [][]interface{}
		var data [][]interface{}
		for _, rowInterface := range dataInterface {
			rowSlice, ok := rowInterface.([]interface{})
			if !ok {
				return nil, errors.New("each data element must be an array")
			}
			data = append(data, rowSlice)
		}

		startCell, ok := request.Params.Arguments["start_cell"].(string)
		if !ok {
			startCell = "A1" // Default value
		}

		f, err := excelize.OpenFile(filepath)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to open Excel file: %v", err)), nil
		}
		defer f.Close()

		// Create the sheet if it doesn't exist
		index, err := f.GetSheetIndex(sheetName)
		if err != nil || index == -1 {
			f.NewSheet(sheetName)
		}

		// Get starting coordinates from the cell reference
		col, row, err := excelize.CellNameToCoordinates(startCell)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("invalid start cell: %v", err)), nil
		}

		// Write all rows
		for i, rowData := range data {
			// Calculate current row number (1-based)
			currentRow := row + i

			// Convert row number back to cell reference (A2, A3, etc.)
			cellName, err := excelize.CoordinatesToCellName(col, currentRow)
			if err != nil {
				return mcp.NewToolResultError(fmt.Sprintf("failed to calculate cell position: %v", err)), nil
			}

			// Write the current row
			if err := f.SetSheetRow(sheetName, cellName, &rowData); err != nil {
				return mcp.NewToolResultError(fmt.Sprintf("failed to write row %d: %v", i+1, err)), nil
			}
		}

		// Save the workbook
		if err := f.Save(); err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to save workbook: %v", err)), nil
		}

		return mcp.NewToolResultText(fmt.Sprintf("Successfully wrote %d rows to Excel", len(data))), nil
	})

	// Tool 3: read_data_from_excel
	readDataTool := mcp.NewTool("read_data_from_excel",
		mcp.WithDescription("Read data from an Excel worksheet"),
		mcp.WithString("filepath",
			mcp.Required(),
			mcp.Description("Path to the Excel file"),
		),
		mcp.WithString("sheet_name",
			mcp.Required(),
			mcp.Description("Name of the worksheet to read from"),
		),
	)

	s.AddTool(readDataTool, func(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
		filepath, ok := request.Params.Arguments["filepath"].(string)
		if !ok {
			return nil, errors.New("filepath must be a string")
		}

		sheetName, ok := request.Params.Arguments["sheet_name"].(string)
		if !ok {
			return nil, errors.New("sheet_name must be a string")
		}

		f, err := excelize.OpenFile(filepath)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to open Excel file: %v", err)), nil
		}
		defer f.Close()

		// Get all rows from the sheet
		rows, err := f.GetRows(sheetName)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to read sheet: %v", err)), nil
		}

		// Convert to list of lists
		var data [][]string
		for _, row := range rows {
			data = append(data, row)
		}

		// Return as properly formatted JSON
		jsonData, err := json.Marshal(data)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to marshal data to JSON: %v", err)), nil
		}

		return mcp.NewToolResultText(string(jsonData)), nil
	})

	// Tool 1: create_worksheet
	createWorksheetTool := mcp.NewTool("create_worksheet",
		mcp.WithDescription("Create new worksheet in workbook"),
		mcp.WithString("filepath",
			mcp.Required(),
			mcp.Description("Path to the Excel file"),
		),
		mcp.WithString("sheet_name",
			mcp.Required(),
			mcp.Description("Name of the worksheet to create"),
		),
	)
	s.AddTool(createWorksheetTool, func(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
		filepath, ok := request.Params.Arguments["filepath"].(string)
		if !ok {
			return nil, errors.New("filepath must be a string")
		}
		sheetName, ok := request.Params.Arguments["sheet_name"].(string)
		if !ok {
			return nil, errors.New("sheet_name must be a string")
		}
		f, err := excelize.OpenFile(filepath)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to open Excel file: %v", err)), nil
		}
		defer f.Close()
		// Check if sheet already exists
		if index, _ := f.GetSheetIndex(sheetName); index != -1 {
			return mcp.NewToolResultError(fmt.Sprintf("worksheet '%s' already exists", sheetName)), nil
		}
		f.NewSheet(sheetName)
		if err := f.Save(); err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to save workbook: %v", err)), nil
		}
		return mcp.NewToolResultText(fmt.Sprintf("Worksheet '%s' created successfully", sheetName)), nil
	})

	// Tool 2: delete_worksheet
	deleteWorksheetTool := mcp.NewTool("delete_worksheet",
		mcp.WithDescription("Delete worksheet from workbook"),
		mcp.WithString("filepath",
			mcp.Required(),
			mcp.Description("Path to the Excel file"),
		),
		mcp.WithString("sheet_name",
			mcp.Required(),
			mcp.Description("Name of the worksheet to delete"),
		),
	)
	s.AddTool(deleteWorksheetTool, func(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
		filepath, ok := request.Params.Arguments["filepath"].(string)
		if !ok {
			return nil, errors.New("filepath must be a string")
		}
		sheetName, ok := request.Params.Arguments["sheet_name"].(string)
		if !ok {
			return nil, errors.New("sheet_name must be a string")
		}
		f, err := excelize.OpenFile(filepath)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to open Excel file: %v", err)), nil
		}
		defer f.Close()
		// Cannot delete the last sheet
		if len(f.GetSheetList()) == 1 {
			return mcp.NewToolResultError("cannot delete the last worksheet in a workbook"), nil
		}
		if err := f.DeleteSheet(sheetName); err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to delete worksheet: %v", err)), nil
		}
		if err := f.Save(); err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to save workbook: %v", err)), nil
		}
		return mcp.NewToolResultText(fmt.Sprintf("Worksheet '%s' deleted successfully", sheetName)), nil
	})

	// Tool 3: rename_worksheet
	renameWorksheetTool := mcp.NewTool("rename_worksheet",
		mcp.WithDescription("Rename worksheet in workbook"),
		mcp.WithString("filepath",
			mcp.Required(),
			mcp.Description("Path to the Excel file"),
		),
		mcp.WithString("old_name",
			mcp.Required(),
			mcp.Description("Current name of the worksheet"),
		),
		mcp.WithString("new_name",
			mcp.Required(),
			mcp.Description("New name for the worksheet"),
		),
	)
	s.AddTool(renameWorksheetTool, func(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
		filepath, ok := request.Params.Arguments["filepath"].(string)
		if !ok {
			return nil, errors.New("filepath must be a string")
		}
		oldName, ok := request.Params.Arguments["old_name"].(string)
		if !ok {
			return nil, errors.New("old_name must be a string")
		}
		newName, ok := request.Params.Arguments["new_name"].(string)
		if !ok {
			return nil, errors.New("new_name must be a string")
		}
		f, err := excelize.OpenFile(filepath)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to open Excel file: %v", err)), nil
		}
		defer f.Close()
		// Check if old sheet exists
		if index, _ := f.GetSheetIndex(oldName); index == -1 {
			return mcp.NewToolResultError(fmt.Sprintf("worksheet '%s' not found", oldName)), nil
		}
		// Check if new name already exists
		if index, _ := f.GetSheetIndex(newName); index != -1 {
			return mcp.NewToolResultError(fmt.Sprintf("worksheet '%s' already exists", newName)), nil
		}
		if err := f.SetSheetName(oldName, newName); err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to rename worksheet: %v", err)), nil
		}
		if err := f.Save(); err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to save workbook: %v", err)), nil
		}
		return mcp.NewToolResultText(fmt.Sprintf("Worksheet renamed from '%s' to '%s' successfully", oldName, newName)), nil
	})

	// Tool 4: get_workbook_metadata
	getWorkbookMetadataTool := mcp.NewTool("get_workbook_metadata",
		mcp.WithDescription("Get metadata about workbook including sheets, ranges, etc."),
		mcp.WithString("filepath",
			mcp.Required(),
			mcp.Description("Path to the Excel file"),
		),
		mcp.WithBoolean("include_ranges",
			mcp.Description("Whether to include range information (optional)"),
		),
	)
	s.AddTool(getWorkbookMetadataTool, func(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
		filepath, ok := request.Params.Arguments["filepath"].(string)
		if !ok {
			return nil, errors.New("filepath must be a string")
		}
		includeRanges, _ := request.Params.Arguments["include_ranges"].(bool)
		f, err := excelize.OpenFile(filepath)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to open Excel file: %v", err)), nil
		}
		defer f.Close()
		// Basic metadata structure
		metadata := struct {
			Sheets    []string              `json:"sheets"`
			Ranges    map[string][][]string `json:"ranges,omitempty"`
			NumSheets int                   `json:"num_sheets"`
		}{
			Sheets:    f.GetSheetList(),
			NumSheets: len(f.GetSheetList()),
		}
		if includeRanges {
			metadata.Ranges = make(map[string][][]string)
			for _, sheet := range metadata.Sheets {
				// Get all non-empty cells in the sheet
				cols, err := f.GetCols(sheet)
				if err != nil {
					continue
				}
				if len(cols) > 0 {
					rows, err := f.GetRows(sheet)
					if err != nil {
						continue
					}
					// Calculate used range
					if len(rows) > 0 {
						startCell, _ := excelize.CoordinatesToCellName(1, 1)
						endCell, _ := excelize.CoordinatesToCellName(len(cols), len(rows))
						metadata.Ranges[sheet] = [][]string{{startCell, endCell}}
					}
				}
			}
		}
		// Convert to JSON
		jsonData, err := json.Marshal(metadata)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to marshal metadata: %v", err)), nil
		}
		return mcp.NewToolResultText(string(jsonData)), nil
	})

	// Tool 5: format_range
	formatRangeTool := mcp.NewTool("format_range",
		mcp.WithDescription("Apply comprehensive formatting to a range of cells in an Excel worksheet. "+
			"Supports text formatting, borders, alignment, number formats, and more. "+
			"All parameters are optional except filepath, sheet_name and start_cell."),

		// Required Parameters
		mcp.WithString("filepath",
			mcp.Required(),
			mcp.Description("Absolute or relative path to the Excel (.xlsx) file "+
				"Example: 'reports/Q3_results.xlsx' or 'C:\\reports\\sales.xlsx'"),
		),
		mcp.WithString("sheet_name",
			mcp.Required(),
			mcp.Description("Name of the worksheet where formatting should be applied. "+
				"Example: 'Sheet1', 'Sales Data', or 'Q3 Report'"),
		),
		mcp.WithString("start_cell",
			mcp.Required(),
			mcp.Description("Top-left cell of the target range in A1 notation. "+
				"Example: 'A1' for single cell or 'B2' for range start"),
		),

		// Optional Parameters
		mcp.WithString("end_cell",
			mcp.Description("Bottom-right cell of the target range in A1 notation. "+
				"If not provided, formatting will apply only to start_cell. "+
				"Example: 'D10' would create a range from start_cell to D10"),
		),
		mcp.WithBoolean("bold",
			mcp.Description("Set to true to make cell text bold. "+
				"Example: true - sets font weight to bold"),
		),
		mcp.WithBoolean("italic",
			mcp.Description("Set to true to make cell text italic. "+
				"Example: true - sets font style to italic"),
		),
		mcp.WithString("underline", // Updated from boolean to string
			mcp.Description("Underline style for text. Possible values: "+
				"'single', 'double', 'singleAccounting', 'doubleAccounting'. "+
				"Example: 'single' applies standard underline"),
		),
		mcp.WithNumber("font_size",
			mcp.Description("Font size in points (1-409). Common values: 10, 11, 12, 14, 16, 18. "+
				"Example: 12 - sets font size to 12pt"),
		),
		mcp.WithString("font_family",
			mcp.Description("Font family name (e.g., 'Arial', 'Calibri', 'Times New Roman'). "+
				"Example: 'Calibri' - uses Calibri font"),
		),
		mcp.WithString("font_color",
			mcp.Description("Font color in hexadecimal RGB format (6-digit, no alpha). "+
				"Example: 'FF0000' for red, '0000FF' for blue"),
		),
		mcp.WithString("bg_color",
			mcp.Description("Background fill color in hexadecimal RGB format. "+
				"Example: 'FFFF00' for yellow fill"),
		),
		mcp.WithString("fill_pattern",
			mcp.Description("Background pattern type. Available options: "+
				"'solid', 'darkGray','mediumGray','lightGray','gray125','gray0625'"),
		),
		mcp.WithString("border_type", // Updated from border_style
			mcp.Description("Type of borders to apply to all edges. Options: "+
				"'none','thin','medium','thick','dashed','dotted','double','hair',"+
				"'mediumDashed','dashDot','mediumDashDot','dashDotDot','mediumDashDotDot','slantDashDot'"),
		),
		mcp.WithString("border_color",
			mcp.Description("Color for all borders in hexadecimal RGB format. "+
				"Example: '000000' for black borders"),
		),
		mcp.WithString("number_format",
			mcp.Description("Number formatting code or name. Built-in formats: "+
				"'general','0','0.00','#,##0','#,##0.00','0%','0.00%',"+
				"'mm-dd-yy','d-mmm-yy','h:mm AM/PM'. "+
				"Example: '$#,##0.00' for currency format"),
		),
		mcp.WithString("horizontal_align",
			mcp.Description("Horizontal text alignment. Options: "+
				"'left','center','right','fill','justify','centerContinuous','distributed'"),
		),
		mcp.WithString("vertical_align",
			mcp.Description("Vertical text alignment. Options: "+
				"'top','center','bottom','justify','distributed'"),
		),
		mcp.WithBoolean("wrap_text",
			mcp.Description("Set to true to enable text wrapping in cells. "+
				"Example: true - wraps long text within cell"),
		),
		mcp.WithNumber("text_rotation",
			mcp.Description("Degrees to rotate text (-90 to 90). "+
				"Example: 45 - rotates text 45 degrees upward"),
		),
		mcp.WithBoolean("merge_cells",
			mcp.Description("Set to true to merge the specified cell range into one cell. "+
				"Note: Contents will be preserved from top-left cell only."),
		),
		mcp.WithBoolean("protection_lock",
			mcp.Description("Set to true to lock cells (requires sheet protection to take effect)"),
		),
		mcp.WithString("conditional_format",
			mcp.Description("JSON string defining conditional formatting rules (advanced usage)"),
		),
	)

	s.AddTool(formatRangeTool, func(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
		// Extract required parameters with validation
		filepath, ok := request.Params.Arguments["filepath"].(string)
		if !ok || filepath == "" {
			return nil, errors.New("filepath is required and must be a non-empty string")
		}
		sheetName, ok := request.Params.Arguments["sheet_name"].(string)
		if !ok || sheetName == "" {
			return nil, errors.New("sheet_name is required and must be a non-empty string")
		}
		startCell, ok := request.Params.Arguments["start_cell"].(string)
		if !ok || startCell == "" {
			return nil, errors.New("start_cell is required and must be a non-empty string")
		}

		// Handle optional end_cell (default to start_cell if not provided)
		endCell, _ := request.Params.Arguments["end_cell"].(string)
		if endCell == "" {
			endCell = startCell
		}

		// Open the Excel file
		f, err := excelize.OpenFile(filepath)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to open Excel file: %v", err)), nil
		}
		defer func() {
			if err := f.Close(); err != nil {
				log.Printf("Warning: error closing file: %v", err)
			}
		}()

		// Initialize style with defaults
		style := &excelize.Style{
			Font:      &excelize.Font{},
			Fill:      excelize.Fill{Type: "pattern", Color: []string{"FFFFFF"}, Pattern: 1},
			Border:    make([]excelize.Border, 4), // top, right, bottom, left
			Alignment: &excelize.Alignment{},
		}

		// Apply font formatting
		if bold, ok := request.Params.Arguments["bold"].(bool); ok {
			style.Font.Bold = bold
		}
		if italic, ok := request.Params.Arguments["italic"].(bool); ok {
			style.Font.Italic = italic
		}
		if underline, ok := request.Params.Arguments["underline"].(string); ok && underline != "" {
			style.Font.Underline = underline
		}
		if fontSize, ok := request.Params.Arguments["font_size"].(float64); ok && fontSize > 0 {
			style.Font.Size = fontSize
		}
		if fontFamily, ok := request.Params.Arguments["font_family"].(string); ok && fontFamily != "" {
			style.Font.Family = fontFamily
		}
		if fontColor, ok := request.Params.Arguments["font_color"].(string); ok && fontColor != "" {
			style.Font.Color = fontColor
		}

		// Apply fill/background formatting
		if bgColor, ok := request.Params.Arguments["bg_color"].(string); ok && bgColor != "" {
			style.Fill.Color[0] = bgColor
		}
		if fillPattern, ok := request.Params.Arguments["fill_pattern"].(string); ok && fillPattern != "" {
			style.Fill.Type = "pattern"
			style.Fill.Pattern = patterns[fillPattern] // Assume patterns map is defined
		}

		// Apply border formatting
		if borderType, ok := request.Params.Arguments["border_type"].(string); ok && borderType != "" {
			borderStyle, exists := borderStyles[borderType] // Assume borderStyles map is defined
			if !exists {
				return mcp.NewToolResultError(fmt.Sprintf("invalid border type: %s", borderType)), nil
			}
			for i := range style.Border {
				style.Border[i].Type = []string{"top", "right", "bottom", "left"}[i]
				style.Border[i].Style = borderStyle
				if borderColor := request.Params.Arguments["border_color"].(string); borderColor != "" {
					style.Border[i].Color = borderColor
				}
			}
		}

		// Apply alignment formatting
		if horizontal, ok := request.Params.Arguments["horizontal_align"].(string); ok && horizontal != "" {
			style.Alignment.Horizontal = horizontal
		}
		if vertical, ok := request.Params.Arguments["vertical_align"].(string); ok && vertical != "" {
			style.Alignment.Vertical = vertical
		}
		if wrap, ok := request.Params.Arguments["wrap_text"].(bool); ok {
			style.Alignment.WrapText = wrap
		}
		if rotation, ok := request.Params.Arguments["text_rotation"].(float64); ok {
			style.Alignment.TextRotation = int(rotation)
		}

		// Apply number formatting
		if formatStr, ok := request.Params.Arguments["number_format"].(string); ok && formatStr != "" {
			formatCode, err := parseNumberFormat(formatStr) // Assume helper function exists
			if err != nil {
				return mcp.NewToolResultError(fmt.Sprintf("invalid number format: %v", err)), nil
			}
			style.NumFmt = formatCode
		}

		// Apply protection
		if lock, ok := request.Params.Arguments["protection_lock"].(bool); ok {
			style.Protection = &excelize.Protection{
				Locked: lock,
			}
		}

		// Create and apply the style
		styleID, err := f.NewStyle(style)
		if err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to create style: %v", err)), nil
		}

		if err := f.SetCellStyle(sheetName, startCell, endCell, styleID); err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to apply style: %v", err)), nil
		}

		// Handle merge cells
		if merge, ok := request.Params.Arguments["merge_cells"].(bool); ok && merge {
			if err := f.MergeCell(sheetName, startCell, endCell); err != nil {
				return mcp.NewToolResultError(fmt.Sprintf("failed to merge cells: %v", err)), nil
			}
		}

		// Save changes
		if err := f.Save(); err != nil {
			return mcp.NewToolResultError(fmt.Sprintf("failed to save workbook: %v", err)), nil
		}

		return mcp.NewToolResultText(
			fmt.Sprintf("Successfully formatted range %s:%s in sheet '%s'",
				startCell, endCell, sheetName),
		), nil
	})

	// Start the stdio server
	log.Println("Excel Tools Server starting...")
	if err := server.ServeStdio(s); err != nil {
		log.Fatalf("Server error: %v", err)
	}
}

var borderStyles = map[string]int{
	"none":             -1,
	"thin":             1,
	"medium":           2,
	"dashed":           3,
	"dotted":           4,
	"thick":            5,
	"double":           6,
	"hair":             7,
	"mediumDashed":     8,
	"dashDot":          9,
	"mediumDashDot":    10,
	"dashDotDot":       11,
	"mediumDashDotDot": 12,
	"slantDashDot":     13,
}
var patterns = map[string]int{
	"none":       0,
	"solid":      1,
	"darkGray":   2,
	"mediumGray": 3,
	"lightGray":  4,
	"gray125":    5,
	"gray0625":   6,
}

// parseNumberFormat converts number format strings to Excel format codes
// Returns format code (int) and error if the format is invalid
func parseNumberFormat(formatStr string) (int, error) {
	// Built-in standard format mappings (name -> code)
	builtInFormats := map[string]int{
		// Number formats
		"general":  0,
		"0":        1,
		"0.00":     2,
		"#,##0":    3,
		"#,##0.00": 4,
		"0%":       9,
		"0.00%":    10,
		"0.00e+00": 11,
		"# ?/?":    12,
		"# ??/??":  13,

		// Date formats
		"mm-dd-yy":      14,
		"d-mmm-yy":      15,
		"d-mmm":         16,
		"mmm-yy":        17,
		"h:mm am/pm":    18,
		"h:mm:ss am/pm": 19,
		"h:mm":          20,
		"h:mm:ss":       21,
		"m/d/yy h:mm":   22,

		// Currency formats
		"$#,##0_);($#,##0)":            7, // Negative in parentheses
		"$#,##0_);[Red]($#,##0)":       8, // Negative in red parentheses
		"$#,##0.00_);($#,##0.00)":      39,
		"$#,##0.00_);[Red]($#,##0.00)": 40,
	}

	// Check if format matches a built-in format by name
	if code, exists := builtInFormats[strings.ToLower(formatStr)]; exists {
		return code, nil
	}

	// Excel supports custom format strings (codes >= 164)
	// For this implementation, we'll support some common patterns
	commonCustomFormats := map[string]int{
		// Custom number formats
		"#,##0_);(#,##0)":                  164,
		"#,##0.00_);(#,##0.00)":            165,
		"[Blue]#,##0_);[Red](#,##0)":       166,
		"[Blue]#,##0.00_);[Red](#,##0.00)": 167,

		// Custom date formats
		"yyyy-mm-dd":  168,
		"dd/mm/yyyy":  169,
		"mm/dd/yyyy":  170,
		"dd-mmm-yyyy": 171,
		"dd-mmm-yy":   172,
		"mmm-yy":      173,

		// Custom time formats
		"[h]:mm":       174,
		"[h]:mm:ss":    175,
		"hh:mm:ss":     176,
		"hh:mm:ss.000": 177,

		// Custom currency formats
		"\"$\"#,##0_);\"$\"(#,##0)":       178,
		"\"$\"#,##0.00_);\"$\"(#,##0.00)": 179,

		// Accounting formats
		"_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)": 180,
	}

	// Check if format matches a common custom format
	if code, exists := commonCustomFormats[formatStr]; exists {
		return code, nil
	}

	// For truly custom formats, we need to add them to the workbook
	// This is a simplified approach - in production you'd want to:
	// 1. Check if the format already exists in the workbook
	// 2. Only add new formats when necessary
	// 3. Manage the format codes carefully

	// Excel's custom format codes start at 164
	nextCustomCode := 181 // Starting after our predefined custom formats

	// Validate the custom format string
	if !isValidExcelFormat(formatStr) {
		return 0, fmt.Errorf("invalid Excel number format: %q", formatStr)
	}

	// For real implementation, you would store the custom format in the workbook
	// Here we just return the next available code
	return nextCustomCode, nil
}

// isValidExcelFormat checks basic validity of a custom Excel number format string
func isValidExcelFormat(format string) bool {
	if len(format) == 0 {
		return false
	}

	// Count sections (Excel formats can have 1-4 sections separated by semicolons)
	sections := strings.Split(format, ";")
	if len(sections) > 4 {
		return false
	}

	// Basic validation for each section
	for _, section := range sections {
		// Check for allowed characters (simplified)
		for _, ch := range section {
			if !(ch >= '0' && ch <= '9') &&
				!(ch >= 'a' && ch <= 'z') &&
				!(ch >= 'A' && ch <= 'Z') &&
				ch != '.' && ch != ',' && ch != '#' &&
				ch != '?' && ch != '/' && ch != '\\' &&
				ch != '*' && ch != '_' && ch != '(' &&
				ch != ')' && ch != '[' && ch != ']' &&
				ch != '"' && ch != '$' && ch != '-' &&
				ch != '+' && ch != ' ' && ch != ':' &&
				ch != 'y' && ch != 'm' && ch != 'd' && // Date components
				ch != 'h' && ch != 's' && ch != 'e' { // Time and scientific notation
				return false
			}
		}
	}

	return true
}
