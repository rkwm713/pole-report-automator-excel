/**
 * Service for processing SPIDA and Katapult data into Excel reports
 */
import * as XLSX from 'xlsx';
import { processKatapultData, formatHeightToString, WireCategory } from '../utils/katapultDataProcessor';

// Types for data structures
export interface PoleData {
  operationNumber: number;
  attachmentAction: string;
  poleOwner: string;
  poleNumber: string;
  poleStructure: string;
  proposedRiser: string;
  proposedGuy: string;
  pla: string;
  constructionGrade: string;
  heightLowestCom: string;
  heightLowestCpsElectrical: string;
  spans: SpanData[];
  fromPole: string;
  toPole: string;
}

export interface SpanData {
  spanHeader: string;
  attachments: AttachmentData[];
}

export interface AttachmentData {
  description: string;
  existingHeight: string;
  proposedHeight: string;
  midSpanProposed: string;
}

export interface ProcessingError {
  code: string;
  message: string;
  details?: string;
}

/**
 * Main class for processing pole data
 */
export class PoleDataProcessor {
  private spidaData: any;
  private katapultData: any;
  private processedPoles: PoleData[] = [];
  private errors: ProcessingError[] = [];
  private poleLookupMap: Map<string, any> = new Map();
  private katapultPoleLookupMap: Map<string, any> = new Map();
  private operationCounter: number = 1;

  constructor() {}

  /**
   * Load and parse the SPIDA JSON data
   */
  loadSpidaData(jsonText: string): boolean {
    try {
      this.spidaData = JSON.parse(jsonText);
      console.log("SPIDA data loaded successfully");
      return true;
    } catch (error) {
      console.error("Failed to parse SPIDA JSON:", error);
      this.errors.push({
        code: "PARSE_ERROR",
        message: "Failed to parse SPIDA JSON file",
        details: error instanceof Error ? error.message : String(error)
      });
      return false;
    }
  }

  /**
   * Load and parse the Katapult JSON data
   */
  loadKatapultData(jsonText: string): boolean {
    try {
      this.katapultData = JSON.parse(jsonText);
      console.log("Katapult data loaded successfully");
      return true;
    } catch (error) {
      console.error("Failed to parse Katapult JSON:", error);
      this.errors.push({
        code: "PARSE_ERROR",
        message: "Failed to parse Katapult JSON file",
        details: error instanceof Error ? error.message : String(error)
      });
      return false;
    }
  }

  /**
   * Check if all required data has been loaded
   */
  isDataLoaded(): boolean {
    return !!this.spidaData && !!this.katapultData;
  }

  /**
   * Get the count of errors
   */
  getErrorCount(): number {
    return this.errors.length;
  }

  /**
   * Get all errors
   */
  getErrors(): ProcessingError[] {
    return this.errors;
  }

  /**
   * Process the loaded data according to the rules
   */
  processData(): boolean {
    if (!this.isDataLoaded()) {
      this.errors.push({
        code: "NO_DATA",
        message: "SPIDA and Katapult data must be loaded before processing"
      });
      return false;
    }

    try {
      // Clear previous processing results
      this.processedPoles = [];
      this.errors = [];
      this.poleLookupMap.clear();
      this.katapultPoleLookupMap.clear();
      this.operationCounter = 1;

      // Add enhanced debug logging for data structure
      console.log("DEBUG: Processing SPIDA and Katapult data");
      console.log("DEBUG: SPIDA data structure:", JSON.stringify(this.spidaData).substring(0, 500) + "...");
      
      if (typeof this.katapultData === 'object') {
        console.log("DEBUG: Katapult data keys:", Object.keys(this.katapultData));
      }
      
      // Create lookup maps for efficient matching between data sources
      this._createPoleLookupMaps();
      
      // Process each pole from SPIDA data
      if (this.spidaData?.leads?.[0]?.locations) {
        console.log(`DEBUG: Found ${this.spidaData.leads[0].locations.length} locations/poles in SPIDA data`);
        
        for (const poleLocationData of this.spidaData.leads[0].locations) {
          try {
            const canonicalPoleId = this._canonicalizePoleID(poleLocationData.label);
            console.log(`DEBUG: Processing pole ${canonicalPoleId} (original label: ${poleLocationData.label})`);
            
            const katapultPoleData = this.katapultPoleLookupMap.get(canonicalPoleId);
            
            if (!katapultPoleData) {
              console.warn(`No matching Katapult data found for pole ${canonicalPoleId}`);
            } else {
              console.log(`DEBUG: Found matching Katapult data for pole ${canonicalPoleId}`);
            }
            
            const poleData = this._extractPoleData(poleLocationData, katapultPoleData);
            if (poleData) {
              // Log the extracted data for debugging
              console.log(`DEBUG: Extracted pole data: ${JSON.stringify({
                poleNumber: poleData.poleNumber,
                poleOwner: poleData.poleOwner,
                poleStructure: poleData.poleStructure,
                pla: poleData.pla,
                constructionGrade: poleData.constructionGrade
              })}`);
              this.processedPoles.push(poleData);
            }
          } catch (error) {
            console.error(`Error processing pole ${poleLocationData?.label}:`, error);
            this.errors.push({
              code: "POLE_PROCESSING_ERROR",
              message: `Error processing pole ${poleLocationData?.label}`,
              details: error instanceof Error ? error.message : String(error)
            });
          }
        }
        
        console.log(`Processed ${this.processedPoles.length} poles`);
        return this.processedPoles.length > 0;
      } else {
        this.errors.push({
          code: "INVALID_STRUCTURE",
          message: "SPIDA data does not contain the expected 'leads[0].locations' structure"
        });
        return false;
      }
    } catch (error) {
      console.error("Error processing pole data:", error);
      this.errors.push({
        code: "PROCESSING_ERROR",
        message: "Error processing pole data",
        details: error instanceof Error ? error.message : String(error)
      });
      return false;
    }
  }

  /**
   * Generate Excel file from processed data
   */
  generateExcel(): Blob | null {
    if (this.processedPoles.length === 0) {
      this.errors.push({
        code: "NO_PROCESSED_DATA",
        message: "No processed pole data available to generate Excel report"
      });
      return null;
    }

    try {
      // Create workbook and worksheet
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([]);
      
      // Add title row
      XLSX.utils.sheet_add_aoa(ws, [["Make Ready Report"]], { origin: "A1" });
      
      // Add header rows with proper merging and styling (dark blue with text wrapping)
      this._addHeaderRows(ws);
      
      // Starting row for data (after headers)
      let currentRow = 4;
      
      // Add data for each pole
      for (const pole of this.processedPoles) {
        const firstRowOfPole = currentRow;
        
        // Initial mapping of data rows
        const endRowsBeforeFromTo = this._calculateEndRow(pole);
        
        // Write pole-level data (columns A-K), including From/To Pole in columns J and K
        this._writePoleData(ws, pole, firstRowOfPole, endRowsBeforeFromTo);
        
        // Write attachment data (columns L-O)
        currentRow = this._writeAttachmentData(ws, pole, firstRowOfPole, endRowsBeforeFromTo);
        
        // Merge cells for pole data (A-I columns)
        this._mergePoleDataCells(ws, firstRowOfPole, endRowsBeforeFromTo);
        
        // Add a blank row between poles for better readability
        currentRow++;
      }
      
      // Add the worksheet to the workbook
      XLSX.utils.book_append_sheet(wb, ws, "Make Ready Report");
      
      // Generate file with style information preserved
      // Note: The XLSX library has limitations with styling, but we've applied the basic style properties
      // For a more comprehensive styling solution, xlsx-style or similar libraries would be needed
      const excelBuffer = XLSX.write(wb, { 
        bookType: 'xlsx', 
        type: 'array',
        cellStyles: true // Attempt to preserve cell styles when possible
      });
      
      return new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    } catch (error) {
      console.error("Error generating Excel:", error);
      this.errors.push({
        code: "EXCEL_GENERATION_ERROR",
        message: "Error generating Excel report",
        details: error instanceof Error ? error.message : String(error)
      });
      return null;
    }
  }
  
  /**
   * PRIVATE: Add header rows to worksheet
   * UPDATED to enhance header formatting based on README specs
   */
  private _addHeaderRows(ws: XLSX.WorkSheet): void {
    // Row 1 (Main Headers)
    XLSX.utils.sheet_add_aoa(ws, [[
      "Operation Number",
      "Attachment Action:\n( I )nstalling\n( R )emoving\n( E )xisting",
      "Pole Owner",
      "Pole #",
      "Pole Structure", 
      "Proposed Riser (Yes/No) &",
      "Proposed Guy (Yes/No) &", 
      "PLA (%) with proposed attachment",
      "Construction Grade of Analysis",
      "Existing Mid-Span Data", // Will be merged J1:K1
      "", 
      "Make Ready Data", // Will be merged L1:O1
      "", "", ""
    ]], { origin: "A1" });
    
    // Row 2 (Sub-Headers)
    XLSX.utils.sheet_add_aoa(ws, [[
      "", "", "", "", "", "", "", "", "",
      "", 
      "",
      "Attachment Height", // Will be merged L2:N2
      "", "",
      "Mid-Span\n(same span as existing)"
    ]], { origin: "A2" });
    
    // Row 3 (Lowest-Level Sub-Headers)
    XLSX.utils.sheet_add_aoa(ws, [[
      "", "", "", "", "", "", "", "", "",
      "Height Lowest Com", 
      "Height Lowest CPS Electrical",
      "Attacher Description",
      "Existing",
      "Proposed",
      "Proposed"
    ]], { origin: "A3" });
    
    // Apply cell merging for headers
    if (!ws['!merges']) ws['!merges'] = [];
    
    // Merge rows 1-3 for columns A-I vertically (as requested)
    for (let col = 0; col < 9; col++) {
      ws['!merges'].push({ s: { r: 0, c: col }, e: { r: 2, c: col } });
    }
    
    // Merge "Existing Mid-Span Data" (J1:K1)
    ws['!merges'].push({ s: { r: 0, c: 9 }, e: { r: 0, c: 10 } });
    
    // Merge "Make Ready Data" (L1:O1)
    ws['!merges'].push({ s: { r: 0, c: 11 }, e: { r: 0, c: 14 } });
    
    // Merge "Attachment Height" (L2:N2)
    ws['!merges'].push({ s: { r: 1, c: 11 }, e: { r: 1, c: 13 } });
    
    // Create cell styling for the worksheet
    if (!ws['!cols']) ws['!cols'] = [];
    
    // Set column widths
    for (let i = 0; i < 15; i++) {
      // Default width for all columns
      ws['!cols'][i] = { wch: 15 };
    }
    
    // Set specific widths for columns that need more space
    ws['!cols'][1] = { wch: 25 }; // Attachment Action column wider
    ws['!cols'][7] = { wch: 20 }; // PLA column
    ws['!cols'][8] = { wch: 20 }; // Construction Grade column
    ws['!cols'][11] = { wch: 30 }; // Attacher Description column wider
    
    // Apply enhanced styling for header rows
    if (!ws['!rows']) ws['!rows'] = [];
    for (let i = 0; i < 3; i++) {
      // Set row heights taller for headers
      ws['!rows'][i] = { hidden: false, hpt: 25 }; // hpt = height in points
    }
    
    // Apply cell styles (dark blue fill with text wrapping for A1-I3)
    if (!ws.A1) ws.A1 = { v: "Operation Number" };
    if (!ws.A1.s) ws.A1.s = {};
    
    // Apply formatting to all header cells in columns A-I, rows 1-3
    const darkBlue = { fgColor: { rgb: "1F4E78" } }; // Dark blue color (Text 2)
    
    // Helper function to get cell reference
    const getCellRef = (col: number, row: number): string => {
      const colLetter = String.fromCharCode(65 + col); // 65 is ASCII for 'A'
      return `${colLetter}${row + 1}`;
    };
    
    // Apply styles to all cells in the A1:I3 range
    for (let col = 0; col < 9; col++) {
      for (let row = 0; row < 3; row++) {
        const cellRef = getCellRef(col, row);
        if (!ws[cellRef]) ws[cellRef] = { v: "" };
        if (!ws[cellRef].s) ws[cellRef].s = {};
        
        // Apply dark blue fill
        ws[cellRef].s.fill = darkBlue;
        
        // Apply text wrapping
        ws[cellRef].s.alignment = { 
          vertical: "center", 
          horizontal: "center", 
          wrapText: true 
        };
        
        // Apply white font color for better contrast
        ws[cellRef].s.font = { 
          color: { rgb: "FFFFFF" }, 
          bold: true 
        };
      }
    }
    
    console.log("DEBUG: Applied dark blue fill and text wrapping to header cells A1:I3");
  }
  
  /**
   * PRIVATE: Calculate end row for pole data including From/To rows
   * FIXED to properly account for From/To Pole rows in the total row count
   */
  private _calculateEndRow(pole: PoleData): number {
    // Count total rows needed for all spans and attachments
    let totalRows = 0;
    
    // For each span group, count header + attachments
    for (const span of pole.spans) {
      // Add 1 for span header
      totalRows++;
      
      // Add rows for each attachment in this span
      totalRows += span.attachments.length;
    }
    
    // Add 2 rows for From/To Pole information
    totalRows += 2;
    
    // If no rows calculated, use at least 3 rows (1 for wire data + 2 for From/To)
    return Math.max(3, totalRows);
  }
  
  /**
   * PRIVATE: Write pole-level data (columns A-K)
   * UPDATED to include From/To Pole data in columns J and K in the correct format
   */
  private _writePoleData(ws: XLSX.WorkSheet, pole: PoleData, row: number, totalRows: number): void {
    // Add pole data for columns A-I
    XLSX.utils.sheet_add_aoa(ws, [[
      pole.operationNumber,
      pole.attachmentAction,
      pole.poleOwner,
      pole.poleNumber,
      pole.poleStructure,
      pole.proposedRiser,
      pole.proposedGuy,
      pole.pla,
      pole.constructionGrade,
      "",  // We'll handle J and K (Height Lowest Com/CPS) separately
      ""
    ]], { origin: `A${row}` });
    
    // Write Height Lowest Com and Height Lowest CPS Electrical in the first row
    XLSX.utils.sheet_add_aoa(ws, [[
      pole.heightLowestCom,
      pole.heightLowestCpsElectrical
    ]], { origin: `J${row}` });
    
    // Calculate the row numbers for From/To pole headers and values
    // The headers should be in the 2nd to last row of the pole section
    // The values should be in the last row of the pole section
    const fromToHeaderRow = row + totalRows - 2;
    const fromToValueRow = row + totalRows - 1;
    
    // Write "From pole" and "To pole" headers in the 2nd to last row
    XLSX.utils.sheet_add_aoa(ws, [[
      "From pole",
      "To pole"
    ]], { origin: `J${fromToHeaderRow}` });
    
    // Write the actual from/to pole values in the last row
    XLSX.utils.sheet_add_aoa(ws, [[
      pole.fromPole,
      pole.toPole
    ]], { origin: `J${fromToValueRow}` });
    
    // Format the header cells
    ["J", "K"].forEach(col => {
      const headerCellRef = `${col}${fromToHeaderRow}`;
      if (!ws[headerCellRef]) ws[headerCellRef] = { v: col === "J" ? "From pole" : "To pole" };
      if (!ws[headerCellRef].s) ws[headerCellRef].s = {};
      
      // Apply bold styling to the header cells
      ws[headerCellRef].s.font = { bold: true };
      ws[headerCellRef].s.alignment = { 
        vertical: "center", 
        horizontal: "center" 
      };
    });
    
    // Log the data being written for debugging
    console.log(`DEBUG: Writing pole data to Excel:`, {
      row,
      operationNumber: pole.operationNumber,
      poleOwner: pole.poleOwner,
      poleNumber: pole.poleNumber,
      poleStructure: pole.poleStructure,
      pla: pole.pla,
      constructionGrade: pole.constructionGrade,
      heightLowestCom: pole.heightLowestCom,
      heightLowestCpsElectrical: pole.heightLowestCpsElectrical,
      fromPole: pole.fromPole,
      toPole: pole.toPole,
      fromToHeaderRow,
      fromToValueRow
    });
  }
  
  /**
   * PRIVATE: Write attachment data (columns L-O)
   * FIXED to properly handle formatting rules and ensure proper row alignment with From/To Pole rows
   */
  private _writeAttachmentData(ws: XLSX.WorkSheet, pole: PoleData, startRow: number, totalRows: number): number {
    let currentRow = startRow;
    
    if (pole.spans.length === 0) {
      // If no spans, write a blank row
      XLSX.utils.sheet_add_aoa(ws, [[
        "", "", "", ""
      ]], { origin: `L${currentRow}` });
      currentRow++;
    } else {
      // For each span group
      for (const span of pole.spans) {
        // Write span header with enhanced formatting
        XLSX.utils.sheet_add_aoa(ws, [[
          span.spanHeader, "", "", ""
        ]], { origin: `L${currentRow}` });
        
        // Apply bold formatting to span header
        const cellRef = `L${currentRow}`;
        if (!ws[cellRef]) ws[cellRef] = { v: span.spanHeader };
        if (!ws[cellRef].s) ws[cellRef].s = {};
        ws[cellRef].s.font = { bold: true };
        
        // Move to next row
        currentRow++;
        
        // Write attachments
        for (const attachment of span.attachments) {
          // Format proposed height according to rules:
          // - If proposed height is different from existing, show it
          // - If same as existing or empty, leave blank
          const proposedHeight = attachment.proposedHeight && 
                                attachment.proposedHeight !== attachment.existingHeight ? 
                                attachment.proposedHeight : "";
          
          // Write attachment data
          XLSX.utils.sheet_add_aoa(ws, [[
            attachment.description,
            attachment.existingHeight,
            proposedHeight,
            attachment.midSpanProposed
          ]], { origin: `L${currentRow}` });
          
          // Log for debugging
          console.log(`DEBUG: Writing attachment: ${attachment.description}, existing: ${attachment.existingHeight}, proposed: ${proposedHeight}, midspan: ${attachment.midSpanProposed}`);
          
          // Move to next row
          currentRow++;
        }
      }
    }
    
    // Calculate rows to skip to align with From/To rows
    // Remember that From/To rows should be the last 2 rows of the pole section
    const expectedEndRow = startRow + totalRows - 2; // -2 because From/To takes 2 rows
    
    // Add blank rows if needed to ensure From/To Pole rows are at the end
    if (currentRow < expectedEndRow) {
      console.log(`DEBUG: Adding ${expectedEndRow - currentRow} blank rows to align with From/To Pole rows`);
      
      // Add blank rows if needed
      while (currentRow < expectedEndRow) {
        XLSX.utils.sheet_add_aoa(ws, [[
          "", "", "", ""
        ]], { origin: `L${currentRow}` });
        currentRow++;
      }
    }
    
    return currentRow;
  }
  
  /**
   * PRIVATE: Merge cells for pole data (A-I columns)
   * UPDATED to only merge columns A-I, since J-K now contain From/To Pole
   */
  private _mergePoleDataCells(ws: XLSX.WorkSheet, startRow: number, rowCount: number): void {
    if (rowCount <= 1) return; // No need to merge if only one row
    
    if (!ws['!merges']) ws['!merges'] = [];
    
    // Merge cells for each column A through I (not J and K anymore)
    for (let col = 0; col < 9; col++) {
      ws['!merges'].push({
        s: { r: startRow - 1, c: col },
        e: { r: startRow + rowCount - 2, c: col }
      });
    }
    
    // Log the merge operations
    console.log(`DEBUG: Merged cells A${startRow}:I${startRow + rowCount - 1}`);
  }
  
  /**
   * PRIVATE: Set column widths
   * UPDATED to adjust J and K column widths for From/To Pole
   */
  private _setColumnWidths(ws: XLSX.WorkSheet): void {
    if (!ws['!cols']) ws['!cols'] = [];
    
    // Set specific widths for each column based on content needs
    const colWidths = [
      15, // A: Operation Number
      20, // B: Attachment Action - wider for wrapped text
      15, // C: Pole Owner
      15, // D: Pole #
      25, // E: Pole Structure (wider for combined info)
      17, // F: Proposed Riser
      17, // G: Proposed Guy
      15, // H: PLA (%)
      20, // I: Construction Grade
      20, // J: From Pole (adjusted width)
      20, // K: To Pole (adjusted width)
      30, // L: Attacher Description (wider for long descriptions)
      15, // M: Existing
      15, // N: Proposed
      20, // O: Mid-Span Proposed
    ];
    
    // Apply column widths
    colWidths.forEach((width, i) => {
      ws['!cols'][i] = { wch: width };
    });
    
    console.log("DEBUG: Set Excel column widths");
  }

  /**
   * Get the processed poles data
   */
  getProcessedPoles(): PoleData[] {
    return this.processedPoles;
  }

  /**
   * Get the count of processed poles
   */
  getProcessedPoleCount(): number {
    return this.processedPoles.length;
  }
  
  /**
   * Test method for generating sample data to verify columns L-O implementation
   * This is useful for debugging and verifying the REF connection logic works correctly
   */
  generateSampleLtoOData(): PoleData {
    // Create a sample pole with attachments to test column L-O logic
    const pole: PoleData = {
      operationNumber: 1,
      attachmentAction: "( E )xisting",
      poleOwner: "CPS",
      poleNumber: "PL410620",
      poleStructure: "40-4 Southern Pine",
      proposedRiser: "NO",
      proposedGuy: "NO",
      pla: "85.2%",
      constructionGrade: "Grade C",
      heightLowestCom: "22'-3\"",
      heightLowestCpsElectrical: "28'-6\"",
      fromPole: "PL410620",
      toPole: "PL398491",
      spans: []
    };
    
    // Add a normal span with a mix of attachments
    pole.spans.push({
      spanHeader: "Ref to PL398491",
      attachments: [
        // Wire with both existing and proposed heights
        {
          description: "CPS Neutral",
          existingHeight: "32'-6\"", 
          proposedHeight: "33'-0\"", // Different from existing
          midSpanProposed: "30'-2\"" // Proposed height available
        },
        // Wire with same existing and proposed (should show blank for proposed)
        {
          description: "CPS Primary",
          existingHeight: "35'-0\"",
          proposedHeight: "35'-0\"", // Same as existing, should be blank in Excel
          midSpanProposed: "33'-4\""
        },
        // Wire with only existing (should have parentheses for midspan)
        {
          description: "AT&T Fiber Optic Com",
          existingHeight: "25'-3\"",
          proposedHeight: "", // No proposed height
          midSpanProposed: "(22'-10\")" // Existing height in parentheses
        },
      ]
    });
    
    // Add a REF connection with special midspan handling
    pole.spans.push({
      spanHeader: "Ref (NE) to service pole",
      attachments: [
        // REF connection with proposed height
        {
          description: "CPS Service",
          existingHeight: "19'-6\"",
          proposedHeight: "20'-0\"",
          midSpanProposed: "18'-0\"" // Proposed height available
        },
        // REF connection with only existing height
        {
          description: "AT&T Com Drop",
          existingHeight: "14'-2\"",
          proposedHeight: "",
          midSpanProposed: "(11'-10\")" // Existing in parentheses
        }
      ]
    });
    
    return pole;
  }
  
  /**
   * Validate that all required data for columns L-O is properly populated
   * This is a utility method to verify that attachment data is complete
   */
  validateAttachmentData(): { valid: boolean, issues: string[] } {
    const issues: string[] = [];
    let valid = true;
    
    if (this.processedPoles.length === 0) {
      return { 
        valid: false, 
        issues: ["No processed poles data to validate"] 
      };
    }
    
    console.log("Validating attachment data for all processed poles...");
    
    // Check all poles and their spans
    for (const pole of this.processedPoles) {
      if (pole.spans.length === 0) {
        issues.push(`Pole ${pole.poleNumber} has no spans defined`);
        valid = false;
        continue;
      }
      
      // Check each span
      for (const span of pole.spans) {
        // Verify span header
        if (!span.spanHeader) {
          issues.push(`Pole ${pole.poleNumber} has a span with no header`);
          valid = false;
        }
        
        // Check for at least one attachment per span
        if (span.attachments.length === 0) {
          issues.push(`Pole ${pole.poleNumber}, span "${span.spanHeader}" has no attachments`);
          valid = false;
          continue;
        }
        
        // Check each attachment
        for (const attachment of span.attachments) {
          // Verify Attacher Description (Column L)
          if (!attachment.description) {
            issues.push(`Pole ${pole.poleNumber}, span "${span.spanHeader}" has an attachment with no description`);
            valid = false;
          }
          
          // Verify existing height has a value (Column M)
          if (!attachment.existingHeight || attachment.existingHeight === "N/A") {
            // This is an info message rather than error, since this could be valid in some cases
            console.log(`Note: Pole ${pole.poleNumber}, span "${span.spanHeader}", attachment "${attachment.description}" has no existing height`);
          }
          
          // Verify mid-span proposed formatting follows rules (Column O)
          if (attachment.midSpanProposed && 
              attachment.midSpanProposed.startsWith("(") && 
              !attachment.midSpanProposed.endsWith(")")) {
            issues.push(`Pole ${pole.poleNumber}, span "${span.spanHeader}", attachment "${attachment.description}" has invalid parenthesis formatting in mid-span proposed height`);
            valid = false;
          }
        }
      }
    }
    
    // Report results
    if (valid) {
      console.log("Attachment data validation passed successfully.");
    } else {
      console.warn(`Attachment data validation failed with ${issues.length} issues.`);
      for (const issue of issues) {
        console.warn(`- ${issue}`);
      }
    }
    
    return { valid, issues };
  }

  /**
   * PRIVATE: Create lookup maps for efficient pole matching
   */
  private _createPoleLookupMaps(): void {
    // Create maps for SPIDA poles
    if (this.spidaData?.leads?.[0]?.locations) {
      for (const poleLocationData of this.spidaData.leads[0].locations) {
        try {
          const canonicalPoleId = this._canonicalizePoleID(poleLocationData.label);
          this.poleLookupMap.set(canonicalPoleId, poleLocationData);
        } catch (error) {
          console.warn(`Could not process pole ${poleLocationData?.label}:`, error);
        }
      }
      console.log(`DEBUG: Created lookup map for ${this.poleLookupMap.size} SPIDA poles`);
    }
    
    // Create maps for Katapult poles - More robust implementation
    try {
      // Log Katapult structure for debugging
      console.log("DEBUG: Katapult structure check at _createPoleLookupMaps");
      
      if (!this.katapultData) {
        console.warn("Katapult data is null or undefined");
        return;
      }
      
      // Try different possible structures for Katapult data
      let katapultNodes: any[] = [];
      
      // Check if nodes exist directly
      if (Array.isArray(this.katapultData.nodes)) {
        console.log("DEBUG: Found nodes array in katapultData.nodes");
        katapultNodes = this.katapultData.nodes;
      } 
      // Check if data is in a 'data' property
      else if (this.katapultData.data && Array.isArray(this.katapultData.data.nodes)) {
        console.log("DEBUG: Found nodes array in katapultData.data.nodes");
        katapultNodes = this.katapultData.data.nodes;
      }
      // Check if data is in a 'features' property (common in GeoJSON)
      else if (Array.isArray(this.katapultData.features)) {
        console.log("DEBUG: Found features array in katapultData.features");
        katapultNodes = this.katapultData.features;
      }
      // Check if the data itself is an array
      else if (Array.isArray(this.katapultData)) {
        console.log("DEBUG: Katapult data itself is an array");
        katapultNodes = this.katapultData;
      }
      // If we still don't have nodes, try to extract from JSON
      else {
        console.log("DEBUG: Attempting to find poles in katapultData keys:", Object.keys(this.katapultData));
        
        // Search for pole data in any property that might contain arrays
        for (const key of Object.keys(this.katapultData)) {
          const value = this.katapultData[key];
          if (Array.isArray(value) && value.length > 0) {
            console.log(`DEBUG: Found array in katapultData.${key} with ${value.length} items`);
            // Check if items have PoleNumber property
            const hasPoleNumberProperty = value.some(item => 
              item && 
              (item.properties?.PoleNumber || 
               item.PoleNumber || 
               item.poleNumber || 
               item.polenumber || 
               item.pole_number)
            );
            
            if (hasPoleNumberProperty) {
              console.log(`DEBUG: Array in katapultData.${key} has items with pole number properties`);
              katapultNodes = value;
              break;
            }
          }
        }
      }
      
      console.log(`DEBUG: Found ${katapultNodes.length} potential pole nodes in Katapult data`);
      
      // Process the nodes
      for (const node of katapultNodes) {
        try {
          if (!node) continue;
          
          // Try various properties where pole number might be stored
          const poleNumber = 
            node.properties?.PoleNumber || 
            node.properties?.poleNumber || 
            node.PoleNumber || 
            node.poleNumber || 
            node.properties?.polenumber || 
            node.polenumber ||
            node.properties?.pole_number || 
            node.pole_number;
          
          if (poleNumber) {
            const canonicalPoleId = this._canonicalizePoleID(poleNumber);
            console.log(`DEBUG: Adding Katapult pole ${canonicalPoleId} to lookup map`);
            this.katapultPoleLookupMap.set(canonicalPoleId, node);
          }
        } catch (error) {
          console.warn(`Could not process Katapult node:`, error);
        }
      }
      
      console.log(`DEBUG: Created Katapult lookup map with ${this.katapultPoleLookupMap.size} poles`);
    } catch (error) {
      console.error("Error creating Katapult lookup map:", error);
    }
    
    console.log(`Created lookup maps: ${this.poleLookupMap.size} SPIDA poles, ${this.katapultPoleLookupMap.size} Katapult poles`);
  }

  /**
   * PRIVATE: Check if any attachment on the pole has changes between measured and recommended designs
   * Used to determine if Column O should be populated
   */
  private _hasPoleAttachmentChanges(poleLocationData: any): boolean {
    try {
      // Find design indices
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) {
        return false;
      }
      
      const measuredDesign = poleLocationData?.designs?.[designIndices.measured]?.structure;
      const recommendedDesign = poleLocationData?.designs?.[designIndices.recommended]?.structure;
      
      if (!measuredDesign || !recommendedDesign) {
        return false;
      }
      
      console.log(`DEBUG: Checking for pole attachment changes`);
      
      // Compare wires between designs
      if (measuredDesign.wires && recommendedDesign.wires) {
        // Create maps for easier comparison
        const measuredWireMap = new Map();
        const recommendedWireMap = new Map();
        
        // Map measured wires
        for (const wire of measuredDesign.wires) {
          const wireKey = `${wire.owner?.id || 'unknown'}-${wire.clientItem?.description || wire.clientItem?.type || 'unknown'}-${wire.attachmentHeight?.value || 0}`;
          measuredWireMap.set(wireKey, wire);
        }
        
        // Check each recommended wire against measured
        for (const wire of recommendedDesign.wires) {
          const wireKey = `${wire.owner?.id || 'unknown'}-${wire.clientItem?.description || wire.clientItem?.type || 'unknown'}-${wire.attachmentHeight?.value || 0}`;
          
          // If wire exists in measured but with different height, it's changed
          const measuredWire = measuredWireMap.get(wireKey);
          if (!measuredWire) {
            console.log(`DEBUG: Found new or moved wire in recommended design: ${wireKey}`);
            return true;
          }
        }
        
        // Check if any measured wires are missing in recommended (removed)
        for (const [wireKey, _] of measuredWireMap.entries()) {
          const wireExists = recommendedDesign.wires.some(w => {
            const recWireKey = `${w.owner?.id || 'unknown'}-${w.clientItem?.description || w.clientItem?.type || 'unknown'}-${w.attachmentHeight?.value || 0}`;
            return recWireKey === wireKey;
          });
          
          if (!wireExists) {
            console.log(`DEBUG: Found removed wire from measured design: ${wireKey}`);
            return true;
          }
        }
      }
      
      // Compare equipment between designs
      if (measuredDesign.equipments && recommendedDesign.equipments) {
        // Create maps for easier comparison
        const measuredEquipMap = new Map();
        const recommendedEquipMap = new Map();
        
        // Map measured equipment
        for (const equip of measuredDesign.equipments) {
          const equipKey = `${equip.owner?.id || 'unknown'}-${equip.clientItem?.type || 'unknown'}-${equip.attachmentHeight?.value || 0}`;
          measuredEquipMap.set(equipKey, equip);
        }
        
        // Check each recommended equipment against measured
        for (const equip of recommendedDesign.equipments) {
          const equipKey = `${equip.owner?.id || 'unknown'}-${equip.clientItem?.type || 'unknown'}-${equip.attachmentHeight?.value || 0}`;
          
          // If equipment exists in measured but with different height, it's changed
          const measuredEquip = measuredEquipMap.get(equipKey);
          if (!measuredEquip) {
            console.log(`DEBUG: Found new or moved equipment in recommended design: ${equipKey}`);
            return true;
          }
        }
        
        // Check if any measured equipment is missing in recommended (removed)
        for (const [equipKey, _] of measuredEquipMap.entries()) {
          const equipExists = recommendedDesign.equipments.some(e => {
            const recEquipKey = `${e.owner?.id || 'unknown'}-${e.clientItem?.type || 'unknown'}-${e.attachmentHeight?.value || 0}`;
            return recEquipKey === equipKey;
          });
          
          if (!equipExists) {
            console.log(`DEBUG: Found removed equipment from measured design: ${equipKey}`);
            return true;
          }
        }
      }
      
      console.log(`DEBUG: No attachment changes found for this pole`);
      return false;
    } catch (error) {
      console.error("Error checking for pole attachment changes:", error);
      return false;
    }
  }
  
  /**
   * PRIVATE: Extract data for a single pole
   */
  private _extractPoleData(poleLocationData: any, katapultPoleData: any): PoleData | null {
    try {
      // Log the full pole data structure for debugging (truncated to avoid massive logs)
      console.log(`DEBUG: Pole location data structure: ${JSON.stringify(poleLocationData).substring(0, 500)}...`);
      
      // Find design indices
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) {
        throw new Error("Could not find required designs in pole data");
      }
      
      console.log(`DEBUG: Found design indices: measured=${designIndices.measured}, recommended=${designIndices.recommended}`);
      
      // Extract basic pole information
      const canonicalPoleId = this._canonicalizePoleID(poleLocationData.label);
      
      // Extract pole owner (prioritize Katapult if available)
      // Updated pole owner extraction with better path traversal
      let poleOwner = this._extractPoleOwner(poleLocationData, katapultPoleData);
      console.log(`DEBUG: Extracted pole owner: ${poleOwner}`);
      
      // Extract pole structure details with improved extraction
      const poleStructure = this._extractPoleStructure(poleLocationData);
      console.log(`DEBUG: Extracted pole structure: ${poleStructure}`);
      
      // Extract proposed riser information
      const proposedRiser = this._extractProposedRiserInfo(poleLocationData, designIndices.recommended);
      
      // Extract proposed guy information
      const proposedGuy = this._extractProposedGuyInfo(poleLocationData, designIndices.recommended);
      
      // Extract PLA value with improved analysis case targeting
      const plaInfo = this._extractPLA(poleLocationData);
      console.log(`DEBUG: Extracted PLA: ${plaInfo.pla} (${plaInfo.actual})`);
      
      // Extract construction grade with improved extraction
      const constructionGrade = this._extractConstructionGrade(poleLocationData);
      console.log(`DEBUG: Extracted construction grade: ${constructionGrade}`);
      
      // Check if pole has any attachment changes between measured and recommended designs
      // This determines whether Column O (Mid-Span Proposed) is displayed
      const poleHasAttachmentChanges = this._hasPoleAttachmentChanges(poleLocationData);
      console.log(`DEBUG: Pole ${canonicalPoleId} has attachment changes: ${poleHasAttachmentChanges}`);
      
      // Extract midspan height data
      const midspanHeights = this._extractExistingMidspanData(poleLocationData, katapultPoleData);
      
      // Extract span data with attachments
      const spans = this._extractSpanData(poleLocationData, designIndices, poleHasAttachmentChanges);
      
      // Extract from/to pole information
      const fromToPoles = this._extractFromToPoles(poleLocationData, katapultPoleData);
      
      // Determine attachment action
      const attachmentAction = this._determineAttachmentAction(poleLocationData);
      
      // Create the pole data object
      const poleData: PoleData = {
        operationNumber: this.operationCounter++,
        attachmentAction: attachmentAction,
        poleOwner: poleOwner,
        poleNumber: canonicalPoleId,
        poleStructure: poleStructure,
        proposedRiser: proposedRiser,
        proposedGuy: proposedGuy,
        pla: plaInfo.pla,
        constructionGrade: constructionGrade,
        heightLowestCom: midspanHeights.com,
        heightLowestCpsElectrical: midspanHeights.electrical,
        spans: spans,
        fromPole: fromToPoles.from,
        toPole: fromToPoles.to
      };
      
      return poleData;
    } catch (error) {
      console.error(`Error extracting pole data:`, error);
      this.errors.push({
        code: "DATA_EXTRACTION_ERROR",
        message: `Failed to extract data for pole ${poleLocationData?.label || "unknown"}`,
        details: error instanceof Error ? error.message : String(error)
      });
      return null;
    }
  }

  /**
   * PRIVATE: Extract pole owner information from multiple possible sources
   * UPDATED to always return "CPS" regardless of input data
   */
  private _extractPoleOwner(poleLocationData: any, katapultPoleData: any): string {
    // Always return CPS as the pole owner
    return "CPS";
  }

  /**
   * PRIVATE: Extract pole structure information
   * UPDATED to properly access pole structure through designs
   */
  private _extractPoleStructure(poleLocationData: any): string {
    try {
      console.log("DEBUG: Extracting pole structure details");
      
      // Find the measured design first
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) {
        console.log("DEBUG: No designs found for pole structure extraction");
        return "Unknown";
      }
      
      // Get the pole from the measured design
      const measuredDesign = poleLocationData?.designs?.[designIndices.measured]?.structure;
      if (!measuredDesign || !measuredDesign.pole) {
        console.log("DEBUG: No pole structure data found in measured design");
        return "Unknown";
      }
      
      const pole = measuredDesign.pole;
      
      // Log available data for debugging
      if (pole.clientItem) {
        console.log("DEBUG: Available clientItem keys:", Object.keys(pole.clientItem));
      }
      
      // Get clientItemAlias (e.g., "40-4")
      let clientItemAlias = pole.clientItemAlias || "";
      
      // Get species
      let species = "";
      if (pole.clientItem?.species) {
        species = pole.clientItem.species;
        console.log(`DEBUG: Extracted pole species: ${species}`);
      } else if (this.spidaData?.clientData?.poles && pole.clientItem?.id) {
        // Try to find species in clientData using the clientItem id
        const clientItemId = pole.clientItem.id;
        for (const clientPole of this.spidaData.clientData.poles) {
          if (clientPole.aliases && Array.isArray(clientPole.aliases)) {
            for (const alias of clientPole.aliases) {
              if (alias.id === clientItemId) {
                species = clientPole.species || "";
                console.log(`DEBUG: Found pole species in clientData: ${species}`);
                break;
              }
            }
            if (species) break;
          }
        }
      }
      
      // If we still don't have species, try class of pole
      let classOfPole = "";
      if (pole.clientItem?.classOfPole) {
        classOfPole = pole.clientItem.classOfPole;
      } else if (this.spidaData?.clientData?.poles && pole.clientItem?.id) {
        // Try to find class in clientData
        const clientItemId = pole.clientItem.id;
        for (const clientPole of this.spidaData.clientData.poles) {
          if (clientPole.aliases && Array.isArray(clientPole.aliases)) {
            for (const alias of clientPole.aliases) {
              if (alias.id === clientItemId) {
                classOfPole = clientPole.classOfPole || "";
                console.log(`DEBUG: Found pole class in clientData: ${classOfPole}`);
                break;
              }
            }
            if (classOfPole) break;
          }
        }
      }
      
      // Combine class and species for full structure info
      let structureStr = "";
      if (clientItemAlias && species) {
        structureStr = `${clientItemAlias} ${species}`;
      } else if (classOfPole && species) {
        structureStr = `${classOfPole} ${species}`;
      } else if (clientItemAlias) {
        structureStr = clientItemAlias;
      } else if (species) {
        structureStr = species;
      } else if (classOfPole) {
        structureStr = classOfPole;
      }
      
      console.log(`DEBUG: Final pole structure string: "${structureStr}"`);
      return structureStr || "Unknown";
    } catch (error) {
      console.error("Error extracting pole structure:", error);
      return "Unknown";
    }
  }

  /**
   * PRIVATE: Find indices for "Measured Design" and "Recommended Design"
   */
  private _findDesignIndices(poleLocationData: any): { measured: number, recommended: number } | null {
    if (!poleLocationData?.designs || !Array.isArray(poleLocationData.designs)) {
      console.log("DEBUG: No designs array found in pole location data");
      return null;
    }
    
    console.log(`DEBUG: Searching for design indices in ${poleLocationData.designs.length} designs`);
    
    // Log all design names for debugging
    const designNames = poleLocationData.designs.map((d: any, i: number) => 
      `[${i}]: ${d?.name || 'Unnamed'}`).join(', ');
    console.log(`DEBUG: Available designs: ${designNames}`);
    
    let measuredIdx = -1;
    let recommendedIdx = -1;
    
    for (let i = 0; i < poleLocationData.designs.length; i++) {
      const design = poleLocationData.designs[i];
      if (design?.name?.includes("Measured Design")) {
        measuredIdx = i;
        console.log(`DEBUG: Found Measured Design at index ${i}`);
      }
      if (design?.name?.includes("Recommended Design")) {
        recommendedIdx = i;
        console.log(`DEBUG: Found Recommended Design at index ${i}`);
      }
    }
    
    // Fallback to first two designs if specific names not found
    if (measuredIdx === -1 && poleLocationData.designs.length > 0) {
      measuredIdx = 0;
      console.log(`DEBUG: Fallback: Using first design as Measured Design (index 0)`);
    }
    if (recommendedIdx === -1 && poleLocationData.designs.length > 1) {
      recommendedIdx = 1;
      console.log(`DEBUG: Fallback: Using second design as Recommended Design (index 1)`);
    } else if (recommendedIdx === -1 && measuredIdx === 0) {
      // If only one design, use it for both
      recommendedIdx = 0;
      console.log(`DEBUG: Fallback: Using the only available design for both Measured and Recommended (index 0)`);
    }
    
    if (measuredIdx === -1 || recommendedIdx === -1) {
      console.log("DEBUG: Could not find required designs");
      return null;
    }
    
    return { measured: measuredIdx, recommended: recommendedIdx };
  }

  /**
   * PRIVATE: Extract proposed riser information
   */
  private _extractProposedRiserInfo(poleLocationData: any, recommendedDesignIdx: number): string {
    try {
      const equipments = poleLocationData?.designs?.[recommendedDesignIdx]?.structure?.equipments;
      if (!equipments || !Array.isArray(equipments)) {
        return "NO";
      }
      
      // Count risers
      const riserCount = equipments.filter(
        eq => eq?.clientItem?.type === "RISER"
      ).length;
      
      return riserCount > 0 ? `YES (${riserCount})` : "NO";
    } catch (error) {
      console.warn("Error extracting riser info:", error);
      return "NO";
    }
  }

  /**
   * PRIVATE: Extract proposed guy information
   */
  private _extractProposedGuyInfo(poleLocationData: any, recommendedDesignIdx: number): string {
    try {
      const guys = poleLocationData?.designs?.[recommendedDesignIdx]?.structure?.guys;
      if (!guys || !Array.isArray(guys)) {
        return "NO";
      }
      
      return guys.length > 0 ? `YES (${guys.length})` : "NO";
    } catch (error) {
      console.warn("Error extracting guy info:", error);
      return "NO";
    }
  }

  /**
   * PRIVATE: Extract PLA value
   * UPDATED to correctly find Pole stress analysis results
   */
  private _extractPLA(poleLocationData: any): { pla: string, actual: number } {
    try {
      // Check if we have designs and find the recommended design
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) {
        console.log("DEBUG: No designs found for PLA extraction");
        return { pla: "N/A", actual: 0 };
      }
      
      // Look for analysis in the recommended design first (leads[0].locations[0].designs[1].analysis)
      const recommendedDesign = poleLocationData?.designs?.[designIndices.recommended];
      if (!recommendedDesign) {
        console.log("DEBUG: No recommended design found");
        return { pla: "N/A", actual: 0 };
      }
      
      let analysisObjects = recommendedDesign.analysis;
      
      // If no analysis at design level, check location level
      if (!analysisObjects || !Array.isArray(analysisObjects) || analysisObjects.length === 0) {
        console.log("DEBUG: No analysis at design level, checking location level");
        analysisObjects = poleLocationData?.analysis;
      }
      
      if (!analysisObjects || !Array.isArray(analysisObjects)) {
        console.log("DEBUG: No analysis objects found at any level");
        return { pla: "N/A", actual: 0 };
      }
      
      console.log(`DEBUG: Found ${analysisObjects.length} analysis objects`);
      
      // Log all analysis case names for debugging
      const caseNames = analysisObjects.map((a: any, i: number) => 
        `[${i}]: ${a?.analysisCaseDetails?.name || 'Unnamed'}`).join(', ');
      console.log(`DEBUG: Available analysis cases: ${caseNames}`);
      
      // First try to find "Light - Grade C" analysis
      for (const analysis of analysisObjects) {
        const caseName = analysis?.analysisCaseDetails?.name || "";
        if (caseName.includes("Light - Grade C")) {
          // Found the target analysis case
          console.log(`DEBUG: Found "Light - Grade C" analysis case`);
          
          // Check for results
          const results = analysis.results || [];
          
          // Find the pole stress result specifically
          for (const result of results) {
            if (result.component === "Pole" && result.analysisType === "STRESS") {
              const plaValue = result.actual;
              if (typeof plaValue === "number") {
                console.log(`DEBUG: Found pole stress value: ${plaValue}`);
                return { 
                  pla: `${plaValue.toFixed(2)}%`, 
                  actual: plaValue 
                };
              }
            }
          }
        }
      }
      
      // If specific analysis not found, try any analysis case with stress results
      for (const analysis of analysisObjects) {
        const results = analysis.results || [];
        for (const result of results) {
          if (result.component === "Pole" && result.analysisType === "STRESS") {
            const plaValue = result.actual;
            if (typeof plaValue === "number") {
              console.log(`DEBUG: Found pole stress in alternative analysis: ${plaValue}`);
              return { 
                pla: `${plaValue.toFixed(2)}%`, 
                actual: plaValue 
              };
            }
          }
        }
      }
      
      console.log("DEBUG: No pole stress result found in any analysis");
      return { pla: "N/A", actual: 0 };
    } catch (error) {
      console.error("Error extracting PLA:", error);
      return { pla: "N/A", actual: 0 };
    }
  }

  /**
   * PRIVATE: Extract construction grade
   * UPDATED to always return "Grade C" as specified
   */
  private _extractConstructionGrade(poleLocationData: any): string {
    // Always return "Grade C" as specified
    return "Grade C";
  }

  /**
   * PRIVATE: Extract existing midspan height data from Katapult data
   * UPDATED to find the absolute lowest midspan height across ALL spans connected to a pole
   * This data is used for Columns J & K (Height Lowest Com & Height Lowest CPS Electrical)
   */
  private _extractExistingMidspanData(poleLocationData: any, katapultPoleData: any): { com: string, electrical: string } {
    try {
      // Only use Katapult data as requested - no fallback to SPIDAcalc
      if (!this.katapultData || !katapultPoleData) {
        console.log("DEBUG: No Katapult data available for midspan heights");
        return { com: "N/A", electrical: "N/A" };
      }
      
      console.log("DEBUG: Extracting midspan heights from Katapult data");
      
      // Extract the pole number/ID (needed for logging)
      const poleNumber = this._canonicalizePoleID(poleLocationData.label);
      console.log(`DEBUG: Looking for midspan data for pole ${poleNumber}`);
      
      // Process Katapult data to get connection and wire information
      const processedKatapultData = processKatapultData(this.katapultData);
      
      // Variables to track lowest heights
      let lowestComHeightInches = Number.MAX_VALUE;
      let lowestElectricalHeightInches = Number.MAX_VALUE;
      let foundCom = false;
      let foundElectrical = false;
      
      // Extract the Katapult internal node ID for this pole
      const katapultNodeId = katapultPoleData.id || '';
      
      // Log the Katapult node ID for debugging
      console.log(`DEBUG: Found Katapult node ID: ${katapultNodeId}`);
      
      // Find all connections involving this pole
      const poleConnections = processedKatapultData.connections.filter(connection => 
        connection.fromPoleId === katapultNodeId || 
        connection.toPoleId === katapultNodeId
      );
      
      console.log(`DEBUG: Found ${poleConnections.length} connections for pole ${poleNumber}`);
      
      // Iterate through each connection
      for (const connection of poleConnections) {
        // Focus on aerial connections (most relevant for midspan)
        if (connection.buttonType !== 'aerial_path' && !connection.buttonType.includes('aerial')) {
          console.log(`DEBUG: Skipping non-aerial connection: ${connection.buttonType}`);
          continue;
        }
        
        console.log(`DEBUG: Processing aerial connection ${connection.connectionId}`);
        
        // Process each wire in the connection
        for (const wireId in connection.wires) {
          const wire = connection.wires[wireId];
          
          // Skip wires with no midspan height
          if (wire.lowestExistingMidspanHeight === null) {
            continue;
          }
          
          // Process based on wire category
          if (wire.category === WireCategory.COMMUNICATION) {
            console.log(`DEBUG: Found COM wire (${wire.company} - ${wire.cableType}) with height ${wire.lowestExistingMidspanHeight}`);
            foundCom = true;
            lowestComHeightInches = Math.min(lowestComHeightInches, wire.lowestExistingMidspanHeight);
          } else if (wire.category === WireCategory.CPS_ELECTRICAL) {
            console.log(`DEBUG: Found CPS ELECTRICAL wire (${wire.company} - ${wire.cableType}) with height ${wire.lowestExistingMidspanHeight}`);
            foundElectrical = true;
            lowestElectricalHeightInches = Math.min(lowestElectricalHeightInches, wire.lowestExistingMidspanHeight);
          }
        }
      }
      
      // Format heights
      const comHeight = foundCom ? 
        formatHeightToString(lowestComHeightInches) : 
        "N/A";
      
      const electricalHeight = foundElectrical ? 
        formatHeightToString(lowestElectricalHeightInches) : 
        "N/A";
      
      // Log results
      console.log(`DEBUG: Extracted lowest midspan heights from Katapult - COM: ${comHeight}, CPS Electrical: ${electricalHeight}`);
      
      // Return results from Katapult data only
      return {
        com: comHeight,
        electrical: electricalHeight
      };
    } catch (error) {
      console.warn("Error extracting midspan heights:", error);
      return { com: "N/A", electrical: "N/A" };
    }
  }

/**
 * PRIVATE: Extract span data with attachments
 * UPDATED to integrate Katapult mid-span height data and handle REF connections
 * Also considers whether pole has attachment changes for Column O display
 */
private _extractSpanData(poleLocationData: any, designIndices: { measured: number, recommended: number }, poleHasAttachmentChanges: boolean): SpanData[] {
  try {
    const spans: SpanData[] = [];
    const recommendedDesign = poleLocationData?.designs?.[designIndices.recommended]?.structure;
    const measuredDesign = poleLocationData?.designs?.[designIndices.measured]?.structure;
    
    if (!recommendedDesign || !recommendedDesign.wireEndPoints) {
      return spans;
    }
    
    // Get pole ID for finding Katapult data
    const poleId = this._canonicalizePoleID(poleLocationData.label);
    console.log(`DEBUG: Extracting span data for pole ${poleId}`);
    
    // Create a map of wire IDs to actual wire objects for both designs
    const recommendedWireMap = this._createWireMap(recommendedDesign);
    const measuredWireMap = this._createWireMap(measuredDesign);
    
    // Process each wire end point (span)
    for (const wireEndPoint of recommendedDesign.wireEndPoints) {
      try {
        // Skip if no wires in this span
        if (!wireEndPoint.wires || wireEndPoint.wires.length === 0) {
          continue;
        }
        
        // Generate span header
        const spanHeader = this._generateSpanHeader(wireEndPoint);
        const isRefConnection = this._isRefConnection(wireEndPoint);
        
        // Get destination pole ID for finding Katapult connections
        const toPoleId = wireEndPoint.structureLabel || '';
        console.log(`DEBUG: Processing span from ${poleId} to ${toPoleId}, isRef: ${isRefConnection}`);
        
        // Create new span data object
        const spanData: SpanData = {
          spanHeader: spanHeader,
          attachments: []
        };
        
        // Process each wire in this span
        for (const wireId of wireEndPoint.wires) {
          // Get wire from recommended design
          const recommendedWire = recommendedWireMap.get(wireId);
          if (!recommendedWire) continue;
          
          // Try to find matching wire in measured design
          const measuredWire = this._findMatchingWire(recommendedWire, measuredWireMap);
          
          // Get attacher description
          const description = this._getAttachmentDescription(recommendedWire);
          
          // Get existing height
          const existingHeight = measuredWire ? 
            this._metersToFeetInches(measuredWire.attachmentHeight?.value) : 
            "N/A";
          
          // Get proposed height
          const proposedHeight = recommendedWire.attachmentHeight?.value ? 
            this._metersToFeetInches(recommendedWire.attachmentHeight.value) : 
            ""; // Leave blank if same as existing
            
          // Get midspan height from Katapult (pass poleHasAttachmentChanges flag)
          const midSpanProposed = this._getMidSpanProposedHeight(
            poleId,
            toPoleId,
            description,
            recommendedWire,
            isRefConnection,
            poleHasAttachmentChanges
          );
          
          // Create attachment data
          const attachmentData: AttachmentData = {
            description,
            existingHeight,
            proposedHeight,
            midSpanProposed
          };
          
          spanData.attachments.push(attachmentData);
        }
        
        // Add span data to spans array
        if (spanData.attachments.length > 0) {
          spans.push(spanData);
        }
      } catch (error) {
        console.warn("Error processing span:", error);
      }
    }
    
    return spans;
  } catch (error) {
    console.error("Error extracting span data:", error);
    return [];
  }
}

/**
 * PRIVATE: Check if a wireEndPoint represents a REF connection
 */
private _isRefConnection(wireEndPoint: any): boolean {
  const type = wireEndPoint.type || "";
  return type !== "NEXT_POLE" && type !== "PREVIOUS_POLE";
}

/**
 * PRIVATE: Get the proposed mid-span height from Katapult data
 * Handles both regular spans and REF connections with enhanced logic for column O
 * Updated to strictly use Katapult midspan data and prioritize owner/type matching for main spans
 */
private _getMidSpanProposedHeight(
  fromPoleId: string,
  toPoleId: string,
  attacherDescription: string,
  spidaWire: any,
  isRefConnection: boolean,
  poleHasAttachmentChanges: boolean = false
): string {
  try {
    // REMOVED: The condition that was forcing "N/A" for all wires if pole has no attachment changes
    // This change allows mid-span data to be displayed even if the pole attachments themselves haven't changed
    
    // Skip if no Katapult data available
    if (!this.katapultData) {
      console.log(`DEBUG: No Katapult data available for mid-span heights`);
      return "N/A";
    }
    
    // Get processed Katapult data
    const processedKatapultData = processKatapultData(this.katapultData);
    
    console.log(`DEBUG: Finding mid-span height for ${attacherDescription} from ${fromPoleId} to ${toPoleId}`);
    
    // First check if this is an underground connection, which should return "UG" instead of "N/A"
    // This implements the fallback logic from excel_gener_details.txt
    const connections = processedKatapultData.connections.filter(connection => 
      (connection.fromPoleId.includes(fromPoleId) && connection.toPoleId.includes(toPoleId)) ||
      (connection.fromPoleId.includes(toPoleId) && connection.toPoleId.includes(fromPoleId)) ||
      (isRefConnection && (connection.fromPoleId.includes(fromPoleId) || connection.toPoleId.includes(fromPoleId)))
    );
    
    // Check if any of these connections are underground paths
    const hasUndergroundPath = connections.some(connection => 
      connection.buttonType === 'underground_path' || 
      connection.buttonType.includes('underground')
    );
    
    if (hasUndergroundPath) {
      console.log(`DEBUG: Found underground path connection, returning "UG" for mid-span proposed`);
      return "UG";
    }
    
    // Get SPIDAcalc attachment height in inches (for context, not primary matching)
    const spidaAttachHeightMeters = spidaWire.attachmentHeight?.value || 0;
    const spidaAttachHeightInches = this._metersToFeet(spidaAttachHeightMeters) * 12;
    
    // Get SPIDAcalc wire details for matching
    const spidaOwner = spidaWire.owner?.id || "";
    const spidaType = spidaWire.clientItem?.description || 
                     spidaWire.clientItem?.type || 
                     spidaWire.clientItem?.size || "";
    
    console.log(`DEBUG: Searching for Katapult wire matching SPIDAcalc: owner="${spidaOwner}", type="${spidaType}"`);
    
    // Look for matching connection in Katapult data
    let matchingConnection;
    
    // For REF connections, we need special handling
    if (isRefConnection) {
      console.log(`DEBUG: Processing REF connection for ${fromPoleId} (${attacherDescription})`);
      
      // Find connections that include fromPoleId
      let fromPoleConnections = processedKatapultData.connections.filter(connection => 
        (connection.fromPoleId.includes(fromPoleId) || connection.toPoleId.includes(fromPoleId))
      );
      
      console.log(`DEBUG: Found ${fromPoleConnections.length} connections for REF from pole ${fromPoleId}`);
      
      // Try to find a more specific match using the wire description
      const descriptionLower = attacherDescription.toLowerCase();
      
      // If description has "service", try to find service drops first
      const isServiceDrop = descriptionLower.includes("service") || descriptionLower.includes("drop");
      if (isServiceDrop) {
        console.log(`DEBUG: Looking for service/drop connections for "${attacherDescription}"`);
        const serviceConnections = fromPoleConnections.filter(connection => 
          connection.buttonType.includes("service") || 
          connection.buttonType.includes("drop") ||
          connection.buttonType === "ref"
        );
        
        if (serviceConnections.length > 0) {
          fromPoleConnections = serviceConnections;
          console.log(`DEBUG: Narrowed down to ${serviceConnections.length} service connections`);
        }
      }
      
      // First try to find a REF connection (buttonType === 'ref')
      matchingConnection = fromPoleConnections.find(connection => 
        connection.isRefConnection
      );
      
      // If no REF connection found, try service drops or taps
      if (!matchingConnection) {
        matchingConnection = fromPoleConnections.find(connection => 
          connection.buttonType.includes('service') || 
          connection.buttonType.includes('drop')
        );
      }
      
      // If still no match, try any non-aerial connection
      if (!matchingConnection) {
        matchingConnection = fromPoleConnections.find(connection => 
          !connection.buttonType.includes('aerial')
        );
      }
      
      // If still nothing, just take the first connection as a last resort
      if (!matchingConnection && fromPoleConnections.length > 0) {
        matchingConnection = fromPoleConnections[0];
        console.log(`DEBUG: Using first available connection as fallback for REF`);
      }
    } else {
      // For regular spans, look for a direct connection between fromPoleId and toPoleId
      console.log(`DEBUG: Looking for aerial span from ${fromPoleId} to ${toPoleId}`);
      matchingConnection = processedKatapultData.connections.find(connection => 
        (connection.fromPoleId.includes(fromPoleId) && connection.toPoleId.includes(toPoleId)) ||
        (connection.fromPoleId.includes(toPoleId) && connection.toPoleId.includes(fromPoleId))
      );
    }
    
    if (matchingConnection) {
      console.log(`DEBUG: Found matching connection: ${matchingConnection.connectionId} (${matchingConnection.buttonType})`);
      
      // Now find the matching wire in this connection
      const wireMatches = [];
      
      for (const wireId in matchingConnection.wires) {
        const wire = matchingConnection.wires[wireId];
        
        // Get wire details for matching
        const wireOwner = wire.company || "";
        const wireType = wire.cableType || "";
        
        // Calculate a match score
        let matchScore = 0;
        
        // OWNER MATCHING - primary criteria
        // Owner exact match
        if (wireOwner.toLowerCase() === spidaOwner.toLowerCase()) {
          matchScore += 25; // Exact match on owner
        } 
        // Owner partial match
        else if (wireOwner.toLowerCase().includes(spidaOwner.toLowerCase()) || 
                 spidaOwner.toLowerCase().includes(wireOwner.toLowerCase())) {
          matchScore += 15; // Partial match on owner
        }
        
        // TYPE MATCHING - secondary criteria
        // Type exact match
        if (wireType.toLowerCase() === spidaType.toLowerCase()) {
          matchScore += 20; // Exact match on type
        }
        // Type partial match
        else if (wireType.toLowerCase().includes(spidaType.toLowerCase()) || 
                 spidaType.toLowerCase().includes(wireType.toLowerCase())) {
          matchScore += 10; // Partial match on type
        }
        
        // Special handling for common owner/type name variations
        const spidaDescLower = attacherDescription.toLowerCase();
        const wireTypeLower = wireType.toLowerCase();
        
        // Check for specific company matches (AT&T, Charter, etc.)
        if ((spidaDescLower.includes("at&t") && wireOwner.toLowerCase().includes("att")) ||
            (spidaDescLower.includes("att") && wireOwner.toLowerCase().includes("at&t")) ||
            (spidaDescLower.includes("charter") && wireOwner.toLowerCase().includes("spectrum")) ||
            (spidaDescLower.includes("spectrum") && wireOwner.toLowerCase().includes("charter")) ||
            (spidaDescLower.includes("cps") && wireOwner.toLowerCase().includes("cps"))) {
          matchScore += 10;
        }
        
        // Check for fiber, service, etc. type matches
        if ((spidaDescLower.includes("fiber") && wireTypeLower.includes("fiber")) ||
            (spidaDescLower.includes("service") && wireTypeLower.includes("service")) ||
            (spidaDescLower.includes("drop") && wireTypeLower.includes("drop")) ||
            (spidaDescLower.includes("primary") && wireTypeLower.includes("primary")) ||
            (spidaDescLower.includes("neutral") && wireTypeLower.includes("neutral"))) {
          matchScore += 10;
        }
        
        // HEIGHT MATCHING - tertiary criteria, lower weight
        // Only use height as a secondary signal for matching if we have midspan observations
        if (wire.midspanObservations && wire.midspanObservations.length > 0) {
          // Find the closest height observation 
          let bestHeightDiff = Number.MAX_VALUE;
          
          for (const obs of wire.midspanObservations) {
            if (obs.originalHeightInches !== null) {
              const heightDiff = Math.abs(obs.originalHeightInches - spidaAttachHeightInches);
              bestHeightDiff = Math.min(bestHeightDiff, heightDiff);
            }
          }
          
          // Only add height score if reasonably close
          if (bestHeightDiff < 60) { // Within 5 feet (60 inches)
            // Lower weight for height matching - max 5 points
            matchScore += Math.max(0, 5 - (bestHeightDiff / 12));
          }
        }
        
        // Add to matches array if score is above threshold
        if (matchScore >= 10) {
          wireMatches.push({
            wire,
            score: matchScore
          });
        }
      }
      
      // Sort matches by score (highest first)
      wireMatches.sort((a, b) => b.score - a.score);
      
      // Use highest scoring match if available
      if (wireMatches.length > 0) {
        const matchingWire = wireMatches[0].wire;
        console.log(`DEBUG: Found matching wire: ${matchingWire.company} - ${matchingWire.cableType} (match score: ${wireMatches[0].score})`);
        
        // Use ONLY midspan heights from Katapult for Column O (never use poleAttachmentObservations)
        const existingHeight = matchingWire.lowestExistingMidspanHeight;
        const proposedHeight = matchingWire.finalProposedMidspanHeight;
        
        console.log(`DEBUG: Wire midspan heights - existing: ${existingHeight}, proposed: ${proposedHeight}`);
        
        // Check if we have valid midspan height data
        if (existingHeight !== null || proposedHeight !== null) {
          // If there's a proposed height, use that without parentheses
          if (proposedHeight !== null) {
            return formatHeightToString(proposedHeight);
          }
          // If there's only an existing height, show in parentheses
          else if (existingHeight !== null) {
            return `(${formatHeightToString(existingHeight)})`;
          }
        } else {
          console.log(`DEBUG: No midspan observations for matched wire in Katapult`);
          return "N/A"; // No midspan data available in Katapult
        }
      }
    }
    
    // If no match found or no midspan data available
    console.log(`DEBUG: No matching wire found or no midspan height data available in Katapult`);
    return "N/A";
  } catch (error) {
    console.error("Error getting mid-span height:", error);
    return "N/A";
  }
}

  /**
   * PRIVATE: Create a map of wire IDs to wire objects
   */
  private _createWireMap(design: any): Map<string, any> {
    const wireMap = new Map();
    
    if (design?.wires && Array.isArray(design.wires)) {
      for (const wire of design.wires) {
        if (wire.id) {
          wireMap.set(wire.id, wire);
        }
      }
    }
    
    if (design?.equipments && Array.isArray(design.equipments)) {
      for (const equipment of design.equipments) {
        if (equipment.id) {
          wireMap.set(equipment.id, equipment);
        }
      }
    }
    
    return wireMap;
  }

  /**
   * PRIVATE: Generate span header based on wireEndPoint data
   */
  private _generateSpanHeader(wireEndPoint: any): string {
    try {
      const type = wireEndPoint.type || "";
      
      // Handle backspan
      if (type === "PREVIOUS_POLE") {
        return "Backspan";
      }
      
      // Handle other types
      const direction = wireEndPoint.direction !== undefined ? 
        this._getDirection(wireEndPoint.direction) : 
        "";
      
      // Get structure label if available
      const structureLabel = wireEndPoint.structureLabel || "";
      
      if (direction && structureLabel) {
        return `Ref (${direction}) to ${structureLabel}`;
      } else if (direction) {
        return `Ref (${direction})`;
      } else if (structureLabel) {
        return `Ref to ${structureLabel}`;
      } else {
        return "Ref";
      }
    } catch (error) {
      console.warn("Error generating span header:", error);
      return "Unknown Span";
    }
  }

  /**
   * PRIVATE: Find matching wire in measured design
   */
  private _findMatchingWire(recommendedWire: any, measuredWireMap: Map<string, any>): any {
    // If we have an id match, use that
    if (recommendedWire.id && measuredWireMap.has(recommendedWire.id)) {
      return measuredWireMap.get(recommendedWire.id);
    }
    
    // Otherwise try to match based on properties
    const wireType = recommendedWire.clientItem?.type;
    const wireOwner = recommendedWire.owner?.id;
    const wireSize = recommendedWire.clientItem?.size;
    
    if (!wireType && !wireOwner && !wireSize) {
      return null;
    }
    
    // Look for a wire with matching properties
    for (const [_, wire] of measuredWireMap.entries()) {
      if (
        wire.clientItem?.type === wireType &&
        wire.owner?.id === wireOwner &&
        wire.clientItem?.size === wireSize
      ) {
        return wire;
      }
    }
    
    return null;
  }

  /**
   * PRIVATE: Get attachment description
   */
  private _getAttachmentDescription(wire: any): string {
    try {
      const owner = wire.owner?.id || "";
      const description = wire.clientItem?.description || "";
      const size = wire.clientItem?.size || "";
      const type = wire.clientItem?.type || "";
      
      if (description) {
        return `${owner} ${description}`.trim();
      } else if (size) {
        return `${owner} ${size} ${type}`.trim();
      } else {
        return `${owner} ${type}`.trim();
      }
    } catch (error) {
      return "Unknown";
    }
  }

  /**
   * PRIVATE: Calculate mid-span height - LEGACY method, now replaced by _getMidSpanProposedHeight
   * Kept for backward compatibility
   */
  private _calculateMidSpanHeight(wire: any): string {
    try {
      // Log deprecation warning
      console.log("DEBUG: _calculateMidSpanHeight is deprecated, use _getMidSpanProposedHeight instead");
      
      // Call the new method if possible, otherwise fall back to simple calculation
      if (!wire.attachmentHeight?.value) {
        return "N/A";
      }
      
      // Assume mid-span is 5-10% lower than attachment height due to sag
      const attachmentHeight = wire.attachmentHeight.value;
      const sagPercentage = 0.07; // 7% sag as an example
      const midSpanHeight = attachmentHeight * (1 - sagPercentage);
      
      return this._metersToFeetInches(midSpanHeight);
    } catch (error) {
      return "N/A";
    }
  }

  /**
   * PRIVATE: Extract from/to pole information
   * ENHANCED to provide more accurate from/to pole relationships
   */
  private _extractFromToPoles(poleLocationData: any, katapultPoleData: any): { from: string, to: string } {
    try {
      // Current pole ID
      const currentPoleId = this._canonicalizePoleID(poleLocationData.label);
      console.log(`DEBUG: Extracting from/to information for pole ${currentPoleId}`);
      
      // Default to using the current pole as the 'from' pole
      let fromPole = currentPoleId;
      let toPole = "N/A";
      
      // Find design indices
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) {
        console.log("DEBUG: No designs found for from/to extraction");
        return { from: fromPole, to: toPole };
      }
      
      const recommendedDesign = poleLocationData?.designs?.[designIndices.recommended]?.structure;
      if (!recommendedDesign || !recommendedDesign.wireEndPoints) {
        console.log("DEBUG: No wire end points found in recommended design");
        return { from: fromPole, to: toPole };
      }
      
      // First try to find a NEXT_POLE type for the 'to' pole
      const nextPoleEndPoint = recommendedDesign.wireEndPoints.find(
        wep => wep.type === "NEXT_POLE" && wep.structureLabel
      );
      
      if (nextPoleEndPoint) {
        toPole = nextPoleEndPoint.structureLabel;
        console.log(`DEBUG: Found NEXT_POLE connection to ${toPole}`);
        return { from: fromPole, to: toPole };
      }
      
      // If no NEXT_POLE, find the first non-PREVIOUS_POLE with a structureLabel
      const nonPrevEndPoint = recommendedDesign.wireEndPoints.find(
        wep => wep.type !== "PREVIOUS_POLE" && wep.structureLabel
      );
      
      if (nonPrevEndPoint) {
        toPole = nonPrevEndPoint.structureLabel;
        console.log(`DEBUG: Found non-PREVIOUS_POLE connection to ${toPole}`);
        return { from: fromPole, to: toPole };
      }
      
      // If we reach here, check Katapult data for connections
      if (this.katapultData && katapultPoleData) {
        console.log("DEBUG: Checking Katapult data for connections");
        
        // Process Katapult data
        const processedKatapultData = processKatapultData(this.katapultData);
        
        // Get node ID
        const katapultNodeId = katapultPoleData.id || '';
        
        // Find all aerial connections involving this pole
        const aerialConnections = processedKatapultData.connections.filter(connection => 
          (connection.fromPoleId === katapultNodeId || connection.toPoleId === katapultNodeId) &&
          (connection.buttonType.includes('aerial') || connection.buttonType === 'aerial_path')
        );
        
        if (aerialConnections.length > 0) {
          console.log(`DEBUG: Found ${aerialConnections.length} aerial connections in Katapult`);
          
          // Find a connection where this pole is the 'from' pole
          const asFromPoleConnection = aerialConnections.find(conn => conn.fromPoleId === katapultNodeId);
          if (asFromPoleConnection) {
            // Find the corresponding node for the 'to' pole
            for (const [nodeId, node] of this.katapultPoleLookupMap.entries()) {
              if (node.id === asFromPoleConnection.toPoleId) {
                toPole = nodeId;
                console.log(`DEBUG: From Katapult - From: ${currentPoleId}, To: ${toPole}`);
                break;
              }
            }
          } else if (aerialConnections.length > 0) {
            // Otherwise just use the first connection
            const firstConnection = aerialConnections[0];
            const otherPoleId = firstConnection.fromPoleId === katapultNodeId ? 
              firstConnection.toPoleId : firstConnection.fromPoleId;
            
            // Find the corresponding node
            for (const [nodeId, node] of this.katapultPoleLookupMap.entries()) {
              if (node.id === otherPoleId) {
                toPole = nodeId;
                console.log(`DEBUG: From Katapult alternative - From: ${currentPoleId}, To: ${toPole}`);
                break;
              }
            }
          }
        }
      }
      
      // Final check for any REF connection if still no 'to' pole
      if (toPole === "N/A") {
        for (const wireEndPoint of recommendedDesign.wireEndPoints) {
          if (wireEndPoint.structureLabel) {
            toPole = wireEndPoint.structureLabel;
            console.log(`DEBUG: Using fallback to first wireEndPoint with label: ${toPole}`);
            break;
          }
        }
      }
      
      console.log(`DEBUG: Final From/To: ${fromPole} / ${toPole}`);
      return { from: fromPole, to: toPole };
    } catch (error) {
      console.warn("Error extracting from/to poles:", error);
      return { from: this._canonicalizePoleID(poleLocationData.label), to: "N/A" };
    }
  }

  /**
   * PRIVATE: Determine attachment action
   */
  private _determineAttachmentAction(poleLocationData: any): string {
    try {
      // This would involve analyzing the difference between measured and recommended designs
      // For now, use a simple algorithm:
      
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) {
        return "( E )xisting";
      }
      
      const measuredDesign = poleLocationData?.designs?.[designIndices.measured]?.structure;
      const recommendedDesign = poleLocationData?.designs?.[designIndices.recommended]?.structure;
      
      if (!measuredDesign || !recommendedDesign) {
        return "( E )xisting";
      }
      
      // Count attachments in each design
      const measuredCount = 
        (measuredDesign.wires?.length || 0) + 
        (measuredDesign.equipments?.length || 0);
      
      const recommendedCount = 
        (recommendedDesign.wires?.length || 0) + 
        (recommendedDesign.equipments?.length || 0);
      
      if (recommendedCount > measuredCount) {
        return "( I )nstalling";
      } else if (recommendedCount < measuredCount) {
        return "( R )emoving";
      } else {
        return "( E )xisting";
      }
    } catch (error) {
      console.warn("Error determining attachment action:", error);
      return "( E )xisting";
    }
  }

  /**
   * PRIVATE: Standardize pole IDs across data sources
   * UPDATED to be more flexible with pole ID formats
   */
  private _canonicalizePoleID(poleID: string): string {
    if (!poleID) return "Unknown";
    
    console.log(`DEBUG: Canonicalizing pole ID: ${poleID}`);
    
    // Extract core pole ID by removing prefixes, etc.
    // First try to match a pattern like "PL######"
    let match = poleID.match(/[A-Z]{2}\d+/);
    if (match) {
      console.log(`DEBUG: Extracted canonical pole ID: ${match[0]} (using regex pattern)`);
      return match[0];
    }
    
    // If that fails, try to extract numeric portion with 1-2 letter prefix
    match = poleID.match(/[A-Z]{1,2}[0-9]+/);
    if (match) {
      console.log(`DEBUG: Extracted canonical pole ID: ${match[0]} (using alternate pattern)`);
      return match[0];
    }
    
    // If both regex attempts fail, remove common prefixes like "1-" or numbers followed by dash
    const cleanId = poleID.replace(/^\d+-/, '');
    if (cleanId !== poleID) {
      console.log(`DEBUG: Cleaned pole ID by removing prefix: ${cleanId}`);
      return cleanId;
    }
    
    console.log(`DEBUG: Using original pole ID: ${poleID}`);
    return poleID;
  }

  /**
   * PRIVATE: Convert meters to feet
   */
  private _metersToFeet(meters: number): number {
    return meters * 3.28084;
  }

  /**
   * PRIVATE: Convert meters to feet-inches format
   */
  private _metersToFeetInches(meters: number): string {
    if (meters === undefined || meters === null) {
      return "N/A";
    }
    
    const feet = this._metersToFeet(meters);
    const wholeFeet = Math.floor(feet);
    const inches = Math.round((feet - wholeFeet) * 12);
    
    // Handle case where inches == 12
    if (inches === 12) {
      return `${wholeFeet + 1}'-0"`;
    }
    
    return `${wholeFeet}'-${inches}"`;
  }

  /**
   * PRIVATE: Convert direction degrees to cardinal direction
   */
  private _getDirection(degrees: number): string {
    const directions = ["N", "NE", "E", "SE", "S", "SW", "W", "NW"];
    const index = Math.round(((degrees % 360) / 45)) % 8;
    return directions[index];
  }
}
