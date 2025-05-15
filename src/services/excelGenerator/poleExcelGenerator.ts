
/**
 * Service for generating Excel reports from pole data
 */
import * as XLSX from 'xlsx';
import { PoleData, AttachmentData } from '../types/poleTypes';

export class PoleExcelGenerator {
  /**
   * Generate Excel file from processed data
   */
  generateExcel(processedPoles: PoleData[]): Blob | null {
    try {
      if (processedPoles.length === 0) {
        console.error("No processed pole data available to generate Excel report");
        return null;
      }

      // Create workbook and worksheet
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([]);
      
      // Add title row
      XLSX.utils.sheet_add_aoa(ws, [["Make Ready Report"]], { origin: "A1" });
      
      // Add header rows with proper merging
      this._addHeaderRows(ws);
      
      // Starting row for data (after headers)
      let currentRow = 4;
      
      // Add data for each pole
      for (const pole of processedPoles) {
        const firstRowOfPole = currentRow;
        
        // Initial mapping of data rows
        const endRowsBeforeFromTo = this._calculateEndRow(pole);
        
        // Write pole-level data (columns A-K)
        this._writePoleData(ws, pole, firstRowOfPole);
        
        // Write attachment data (columns L-O)
        currentRow = this._writeAttachmentData(ws, pole, firstRowOfPole, endRowsBeforeFromTo);
        
        // Merge cells for pole data (A-K columns)
        this._mergePoleDataCells(ws, firstRowOfPole, endRowsBeforeFromTo);
        
        // Add From/To Pole rows
        currentRow = this._writeFromToPoleData(ws, pole, currentRow);
        
        // Add a blank row between poles for better readability
        currentRow++;
      }
      
      // Set column widths
      this._setColumnWidths(ws);
      
      // Add the worksheet to the workbook
      XLSX.utils.book_append_sheet(wb, ws, "Make Ready Report");
      
      // Generate file
      const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      return new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    } catch (error) {
      console.error("Error generating Excel:", error);
      return null;
    }
  }

  /**
   * PRIVATE: Add header rows to worksheet
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
      "Height Lowest Com", 
      "Height Lowest CPS Electrical",
      "Attachment Height", // Will be merged L2:N2
      "", "",
      "Mid-Span\n(same span as existing)"
    ]], { origin: "A2" });
    
    // Row 3 (Lowest-Level Sub-Headers)
    XLSX.utils.sheet_add_aoa(ws, [[
      "", "", "", "", "", "", "", "", "", "", "",
      "Attacher Description",
      "Existing",
      "Proposed",
      "Proposed"
    ]], { origin: "A3" });
    
    // Apply cell merging for headers
    if (!ws['!merges']) ws['!merges'] = [];
    
    // Merge "Existing Mid-Span Data" (J1:K1)
    ws['!merges'].push({ s: { r: 0, c: 9 }, e: { r: 0, c: 10 } });
    
    // Merge "Make Ready Data" (L1:O1)
    ws['!merges'].push({ s: { r: 0, c: 11 }, e: { r: 0, c: 14 } });
    
    // Merge "Attachment Height" (L2:N2)
    ws['!merges'].push({ s: { r: 1, c: 11 }, e: { r: 1, c: 13 } });
    
    // Apply enhanced styling for header rows
    // Note: XLSX doesn't support extensive styling, 
    // but we can set basic properties

    // Set bold for headers (via XLSX utils limited formatting)
    if (!ws['!rows']) ws['!rows'] = [];
    for (let i = 0; i < 3; i++) {
      // Set row heights a bit taller for headers
      ws['!rows'][i] = { hidden: false, hpt: 25 }; // hpt = height in points
    }
  }
  
  /**
   * PRIVATE: Calculate end row for pole data before From/To rows
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
    
    // If no rows calculated, use at least 1 row for the pole
    return Math.max(1, totalRows);
  }
  
  /**
   * PRIVATE: Write pole-level data (columns A-K)
   */
  private _writePoleData(ws: XLSX.WorkSheet, pole: PoleData, row: number): void {
    // Add pole data
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
      pole.heightLowestCom,
      pole.heightLowestCpsElectrical
    ]], { origin: `A${row}` });
    
    // Log the data being written for debugging
    console.log(`DEBUG: Writing pole data to Excel:`, {
      row,
      operationNumber: pole.operationNumber,
      poleOwner: pole.poleOwner,
      poleNumber: pole.poleNumber,
      poleStructure: pole.poleStructure,
      pla: pole.pla,
      constructionGrade: pole.constructionGrade
    });
  }
  
  /**
   * PRIVATE: Write attachment data (columns L-O)
   */
  private _writeAttachmentData(ws: XLSX.WorkSheet, pole: PoleData, startRow: number, totalRows: number): number {
    let currentRow = startRow;
    
    if (pole.spans.length === 0) {
      // If no spans, write a blank row
      XLSX.utils.sheet_add_aoa(ws, [[
        "", "", "", ""
      ]], { origin: `L${currentRow}` });
      return currentRow + 1;
    }
    
    // For each span group
    for (const span of pole.spans) {
      // Write span header with enhanced formatting
      XLSX.utils.sheet_add_aoa(ws, [[
        span.spanHeader, "", "", ""
      ]], { origin: `L${currentRow}` });
      
      // Move to next row
      currentRow++;
      
      // Write attachments
      for (const attachment of span.attachments) {
        // Write attachment data
        XLSX.utils.sheet_add_aoa(ws, [[
          attachment.description,
          attachment.existingHeight,
          attachment.proposedHeight,
          attachment.midSpanProposed
        ]], { origin: `L${currentRow}` });
        
        // Log for debugging
        console.log(`DEBUG: Writing attachment: ${attachment.description}, existing: ${attachment.existingHeight}, proposed: ${attachment.proposedHeight}`);
        
        // Move to next row
        currentRow++;
      }
    }
    
    return currentRow;
  }
  
  /**
   * PRIVATE: Merge cells for pole data (A-K)
   */
  private _mergePoleDataCells(ws: XLSX.WorkSheet, startRow: number, rowCount: number): void {
    if (rowCount <= 1) return; // No need to merge if only one row
    
    if (!ws['!merges']) ws['!merges'] = [];
    
    // Merge cells for each column A through K
    for (let col = 0; col < 11; col++) {
      ws['!merges'].push({
        s: { r: startRow - 1, c: col },
        e: { r: startRow + rowCount - 2, c: col }
      });
    }
    
    // Log the merge operations
    console.log(`DEBUG: Merged cells A${startRow}:K${startRow + rowCount - 1}`);
  }
  
  /**
   * PRIVATE: Write From/To Pole data
   */
  private _writeFromToPoleData(ws: XLSX.WorkSheet, pole: PoleData, currentRow: number): number {
    // Add "From Pole" row
    XLSX.utils.sheet_add_aoa(ws, [[
      "", "", "", "", "", "", "", "", "", "", "",
      "From Pole", pole.fromPole, "", ""
    ]], { origin: `A${currentRow}` });
    
    console.log(`DEBUG: Writing From Pole: ${pole.fromPole} at row ${currentRow}`);
    
    currentRow++;
    
    // Add "To Pole" row
    XLSX.utils.sheet_add_aoa(ws, [[
      "", "", "", "", "", "", "", "", "", "", "",
      "To Pole", pole.toPole, "", ""
    ]], { origin: `A${currentRow}` });
    
    console.log(`DEBUG: Writing To Pole: ${pole.toPole} at row ${currentRow}`);
    
    return currentRow + 1;
  }
  
  /**
   * PRIVATE: Set column widths
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
      20, // J: Height Lowest Com
      20, // K: Height Lowest CPS Electrical
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
}
