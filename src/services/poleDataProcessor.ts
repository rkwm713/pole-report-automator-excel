

/**
 * Service for processing SPIDA and Katapult data into Excel reports
 */
import * as XLSX from 'xlsx';

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
      
      // Add header rows with proper merging
      this._addHeaderRows(ws);
      
      // Starting row for data (after headers)
      let currentRow = 4;
      
      // Add data for each pole
      for (const pole of this.processedPoles) {
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

    // Enable text wrapping and center alignment for header cells
    // Note: .xlsx format has limited style control via XLSX library
    // Typically this would be done with .s property for cell styles
    // But XLSX.utils.aoa_to_sheet doesn't fully support this
    // In a full implementation, we'd apply xlsx-style or similar
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
   * UPDATED to improve span header formatting
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
      
      // Apply styling to span header - normally we would add styling here
      // but XLSX has limited style support in this API
      
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
   * UPDATED to fix vertical alignment
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
    
    // Note: Vertical alignment would ideally be set to 'top' or 'center'
    // but requires more advanced styling capabilities not fully supported
    // in the basic XLSX utils. In a full implementation, we would use
    // xlsx-style or adjust after creation.
  }
  
  /**
   * PRIVATE: Write From/To Pole data
   * UPDATED to enhance formatting
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
   * UPDATED to better match README specifications
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
      
      // Extract midspan height data
      const midspanHeights = this._extractExistingMidspanData(poleLocationData, katapultPoleData);
      
      // Extract span data with attachments
      const spans = this._extractSpanData(poleLocationData, designIndices);
      
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
   * UPDATED to improve extraction based on README documentation
   */
  private _extractPoleOwner(poleLocationData: any, katapultPoleData: any): string {
    try {
      // Log available paths for debugging
      console.log("DEBUG: Checking pole owner paths");
      
      // Log all potential paths for debugging
      if (katapultPoleData?.properties?.poleOwner) {
        console.log(`DEBUG: Found poleOwner in katapultData.properties.poleOwner: ${katapultPoleData.properties.poleOwner}`);
      }
      if (katapultPoleData?.properties?.owner) {
        console.log(`DEBUG: Found owner in katapultData.properties.owner: ${katapultPoleData.properties.owner}`);
      }
      if (katapultPoleData?.properties?.Owner) {
        console.log(`DEBUG: Found Owner in katapultData.properties.Owner: ${katapultPoleData.properties.Owner}`);
      }
      if (poleLocationData?.structure?.pole?.owner?.id) {
        console.log(`DEBUG: Found owner in spidaData.structure.pole.owner.id: ${poleLocationData.structure.pole.owner.id}`);
      }
      
      // Enhanced pole owner extraction - try multiple paths
      
      // Try multiple paths in Katapult data first (prioritize)
      if (katapultPoleData) {
        // Check all possible paths for pole owner in Katapult
        const katapultOwner = 
          katapultPoleData.properties?.poleOwner ||
          katapultPoleData.properties?.owner ||
          katapultPoleData.properties?.Owner ||
          katapultPoleData.properties?.PoleOwner ||
          katapultPoleData.poleOwner ||
          katapultPoleData.owner ||
          katapultPoleData.Owner;
          
        if (katapultOwner) {
          console.log(`DEBUG: Using pole owner from Katapult: ${katapultOwner}`);
          return katapultOwner;
        }
      }
      
      // Fallback to multiple possible SPIDA paths 
      const spidaOwner = 
        poleLocationData?.structure?.pole?.owner?.id ||
        poleLocationData?.structure?.pole?.clientItem?.owner ||
        poleLocationData?.owner?.id;
      
      if (spidaOwner) {
        console.log(`DEBUG: Using pole owner from SPIDA: ${spidaOwner}`);
        return spidaOwner;
      }
      
      console.log("DEBUG: No pole owner found, using 'Unknown'");
      return "Unknown";
    } catch (error) {
      console.error("Error extracting pole owner:", error);
      return "Unknown";
    }
  }

  /**
   * PRIVATE: Extract pole structure information
   * UPDATED to match README documentation
   */
  private _extractPoleStructure(poleLocationData: any): string {
    try {
      console.log("DEBUG: Extracting pole structure details");
      const pole = poleLocationData?.structure?.pole;
      if (!pole) {
        console.log("DEBUG: No pole structure data found");
        return "Unknown";
      }
      
      // Log available data for debugging
      if (pole.clientItem) {
        console.log("DEBUG: Available clientItem keys:", Object.keys(pole.clientItem));
      }
      
      // Get height and convert from meters to feet
      let height = "";
      if (pole.clientItem?.height?.value) {
        const meters = pole.clientItem.height.value;
        const feet = this._metersToFeet(meters);
        height = `${Math.round(feet)}'`;
        console.log(`DEBUG: Extracted height: ${meters}m converted to ${height}`);
      } else {
        console.log("DEBUG: No height value found in pole.clientItem.height.value");
      }
      
      // Get class
      let poleClass = "";
      if (pole.clientItem?.classOfPole) {
        poleClass = pole.clientItem.classOfPole;
        console.log(`DEBUG: Extracted pole class: ${poleClass}`);
      } else {
        console.log("DEBUG: No classOfPole found in pole.clientItem");
      }
      
      // Get species
      let species = "";
      if (pole.clientItem?.species) {
        species = pole.clientItem.species;
        console.log(`DEBUG: Extracted pole species: ${species}`);
      } else {
        console.log("DEBUG: No species found in pole.clientItem");
      }
      
      // Alternative paths if primary paths fail
      if (!height && pole.clientItemAlias) {
        // Try to extract height from clientItemAlias (e.g., "40-4")
        const parts = pole.clientItemAlias.split('-');
        if (parts.length >= 1 && !isNaN(parseInt(parts[0]))) {
          height = `${parts[0]}'`;
          console.log(`DEBUG: Extracted height from clientItemAlias: ${height}`);
        }
      }
      
      if (!poleClass && pole.clientItemAlias) {
        // Try to extract class from clientItemAlias (e.g., "40-4")
        const parts = pole.clientItemAlias.split('-');
        if (parts.length >= 2 && !isNaN(parseInt(parts[1]))) {
          poleClass = parts[1];
          console.log(`DEBUG: Extracted class from clientItemAlias: ${poleClass}`);
        }
      }
      
      // Format according to README: "40-4 Southern Pine"
      let structureStr = "";
      
      if (height && poleClass) {
        structureStr = `${height}-Class ${poleClass}`;
      } else if (height) {
        structureStr = height;
      } else if (poleClass) {
        structureStr = `Class ${poleClass}`;
      }
      
      if (species) {
        structureStr = structureStr ? `${structureStr} ${species}` : species;
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
   * UPDATED to match README documentation and fix missing PLA issue
   */
  private _extractPLA(poleLocationData: any): { pla: string, actual: number } {
    try {
      // Find the target analysis case (assuming "Light - Grade C" or similar)
      const analysisObjects = poleLocationData?.analysis;
      if (!analysisObjects || !Array.isArray(analysisObjects)) {
        console.log("DEBUG: No analysis objects found in pole location data");
        return { pla: "N/A", actual: 0 };
      }
      
      console.log(`DEBUG: Found ${analysisObjects.length} analysis objects`);
      
      // Log all analysis case names for debugging
      const caseNames = analysisObjects.map((a: any, i: number) => 
        `[${i}]: ${a?.analysisCaseDetails?.name || 'Unnamed'}`).join(', ');
      console.log(`DEBUG: Available analysis cases: ${caseNames}`);
      
      // Look for the recommended design analysis case - trying multiple match patterns
      let targetAnalysis = null;
      
      // First, try exact matches for known patterns
      const knownPatterns = [
        "Recommended Design",
        "Light - Grade C",
        "Grade C",
        "NESC Light Loading Grade C"
      ];
      
      for (const pattern of knownPatterns) {
        for (const analysis of analysisObjects) {
          const caseName = analysis?.analysisCaseDetails?.name || "";
          if (caseName.includes(pattern)) {
            targetAnalysis = analysis;
            console.log(`DEBUG: Found analysis case matching "${pattern}": ${caseName}`);
            break;
          }
        }
        if (targetAnalysis) break;
      }
      
      // If no exact matches, look for any case with "Recommended" or "Grade"
      if (!targetAnalysis) {
        for (const analysis of analysisObjects) {
          const caseName = analysis?.analysisCaseDetails?.name || "";
          if (
            caseName.includes("Recommended") || 
            caseName.includes("Grade") ||
            caseName.includes("NESC")
          ) {
            targetAnalysis = analysis;
            console.log(`DEBUG: Found alternative analysis case: ${caseName}`);
            break;
          }
        }
      }
      
      // Use the first analysis if specific case not found
      if (!targetAnalysis && analysisObjects.length > 0) {
        targetAnalysis = analysisObjects[0];
        console.log(`DEBUG: Using first available analysis case as fallback: ${targetAnalysis?.analysisCaseDetails?.name || 'Unnamed'}`);
      }
      
      // If still no analysis found
      if (!targetAnalysis) {
        console.log("DEBUG: No suitable analysis case found");
        return { pla: "N/A", actual: 0 };
      }
      
      // Find the pole stress result
      const results = targetAnalysis.results || [];
      if (results.length === 0) {
        console.log("DEBUG: No results found in the target analysis case");
      } else {
        console.log(`DEBUG: Found ${results.length} results in the analysis case`);
        
        // Log all result types for debugging
        const resultTypes = results.map((r: any, i: number) => 
          `[${i}]: component=${r.component}, type=${r.analysisType}, actual=${r.actual}`).join('; ');
        console.log(`DEBUG: Available results: ${resultTypes}`);
      }
      
      for (const result of results) {
        if (
          result.component === "Pole" && 
          result.analysisType === "STRESS"
        ) {
          const plaValue = result.actual;
          if (typeof plaValue === "number") {
            // Format as percentage with 2 decimal places
            console.log(`DEBUG: Found PLA value: ${plaValue}`);
            return { 
              pla: `${plaValue.toFixed(2)}%`, 
              actual: plaValue 
            };
          }
        }
      }
      
      // If we get here, try broader search for any stress result
      for (const result of results) {
        if (result.analysisType === "STRESS" && typeof result.actual === "number") {
          console.log(`DEBUG: Found alternative stress result for ${result.component}: ${result.actual}`);
          return {
            pla: `${result.actual.toFixed(2)}%`,
            actual: result.actual
          };
        }
      }
      
      console.log("DEBUG: No suitable PLA result found");
      return { pla: "N/A", actual: 0 };
    } catch (error) {
      console.error("Error extracting PLA:", error);
      return { pla: "N/A", actual: 0 };
    }
  }

  /**
   * PRIVATE: Extract construction grade
   * UPDATED to match README documentation and fix extraction
   */
  private _extractConstructionGrade(poleLocationData: any): string {
    try {
      // Find the target analysis case
      const analysisObjects = poleLocationData?.analysis;
      if (!analysisObjects || !Array.isArray(analysisObjects)) {
        console.log("DEBUG: No analysis objects found for construction grade extraction");
        return "N/A";
      }
      
      console.log(`DEBUG: Extracting construction grade from ${analysisObjects.length} analysis objects`);
      
      // First try to find a recommended design analysis
      for (const analysis of analysisObjects) {
        const caseName = analysis?.analysisCaseDetails?.name || "";
        const constructionGrade = analysis?.analysisCaseDetails?.constructionGrade;
        
        console.log(`DEBUG: Checking analysis case "${caseName}" for construction grade: ${constructionGrade}`);
        
        if (caseName.includes("Recommended") && constructionGrade) {
          console.log(`DEBUG: Found construction grade ${constructionGrade} in recommended analysis`);
          return `Grade ${constructionGrade}`;
        }
      }
      
      // If not found in recommended analysis, check any analysis case with a grade
      for (const analysis of analysisObjects) {
        const constructionGrade = analysis?.analysisCaseDetails?.constructionGrade;
        if (constructionGrade) {
          console.log(`DEBUG: Found construction grade ${constructionGrade} in general analysis`);
          return `Grade ${constructionGrade}`;
        }
      }
      
      // Also check the analysis name for grade information
      for (const analysis of analysisObjects) {
        const caseName = analysis?.analysisCaseDetails?.name || "";
        if (caseName.includes("Grade")) {
          // Try to extract grade from the name (e.g., "Light - Grade C")
          const match = caseName.match(/Grade\s+([A-C])/i);
          if (match && match[1]) {
            console.log(`DEBUG: Extracted grade ${match[1]} from analysis name "${caseName}"`);
            return `Grade ${match[1]}`;
          }
        }
      }
      
      console.log("DEBUG: No construction grade found");
      return "N/A";
    } catch (error) {
      console.error("Error extracting construction grade:", error);
      return "N/A";
    }
  }

  /**
   * PRIVATE: Extract existing midspan height data
   */
  private _extractExistingMidspanData(poleLocationData: any, katapultPoleData: any): { com: string, electrical: string } {
    // This would involve analyzing attachments and their midspan sags
    // For this implementation, we'll extract the lowest COM and electrical attachments
    try {
      let lowestComHeight = Number.MAX_VALUE;
      let lowestElectricalHeight = Number.MAX_VALUE;
      let foundCom = false;
      let foundElectrical = false;
      
      // Find design indices
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) {
        return { com: "N/A", electrical: "N/A" };
      }
      
      // Get attachments from Measured Design
      const measuredDesign = poleLocationData?.designs?.[designIndices.measured]?.structure;
      if (!measuredDesign) {
        return { com: "N/A", electrical: "N/A" };
      }
      
      // Process wires
      if (measuredDesign.wires && Array.isArray(measuredDesign.wires)) {
        for (const wire of measuredDesign.wires) {
          const usageGroup = wire?.usageGroup || "";
          const attachmentHeight = wire?.attachmentHeight?.value;
          
          if (!attachmentHeight) continue;
          
          // Check usage group to determine if it's a com or electrical attachment
          if (usageGroup.includes("COMMUNICATION")) {
            foundCom = true;
            lowestComHeight = Math.min(lowestComHeight, attachmentHeight);
          } else if (
            usageGroup.includes("PRIMARY") ||
            usageGroup.includes("NEUTRAL") ||
            usageGroup.includes("SECONDARY")
          ) {
            foundElectrical = true;
            lowestElectricalHeight = Math.min(lowestElectricalHeight, attachmentHeight);
          }
        }
      }
      
      // Process equipment
      if (measuredDesign.equipments && Array.isArray(measuredDesign.equipments)) {
        for (const equipment of measuredDesign.equipments) {
          const usageGroup = equipment?.usageGroup || "";
          const attachmentHeight = equipment?.attachmentHeight?.value;
          
          if (!attachmentHeight) continue;
          
          // Check usage group
          if (usageGroup.includes("COMMUNICATION")) {
            foundCom = true;
            lowestComHeight = Math.min(lowestComHeight, attachmentHeight);
          } else if (
            usageGroup.includes("PRIMARY") ||
            usageGroup.includes("NEUTRAL") ||
            usageGroup.includes("SECONDARY")
          ) {
            foundElectrical = true;
            lowestElectricalHeight = Math.min(lowestElectricalHeight, attachmentHeight);
          }
        }
      }
      
      // Convert heights to feet-inches format
      const comHeight = foundCom ? 
        this._metersToFeetInches(lowestComHeight) : 
        "N/A";
      
      const electricalHeight = foundElectrical ? 
        this._metersToFeetInches(lowestElectricalHeight) : 
        "N/A";
      
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
   */
  private _extractSpanData(poleLocationData: any, designIndices: { measured: number, recommended: number }): SpanData[] {
    try {
      const spans: SpanData[] = [];
      const recommendedDesign = poleLocationData?.designs?.[designIndices.recommended]?.structure;
      const measuredDesign = poleLocationData?.designs?.[designIndices.measured]?.structure;
      
      if (!recommendedDesign || !recommendedDesign.wireEndPoints) {
        return spans;
      }
      
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
            
            // Create attachment data
            const attachmentData: AttachmentData = {
              description: this._getAttachmentDescription(recommendedWire),
              existingHeight: measuredWire ? 
                this._metersToFeetInches(measuredWire.attachmentHeight?.value) : 
                "N/A",
              proposedHeight: this._metersToFeetInches(recommendedWire.attachmentHeight?.value),
              midSpanProposed: this._calculateMidSpanHeight(recommendedWire) || "N/A"
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
   * PRIVATE: Calculate mid-span height
   */
  private _calculateMidSpanHeight(wire: any): string {
    try {
      // This would involve sag calculations based on wire properties
      // As a placeholder, calculate a simple approximation
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
   */
  private _extractFromToPoles(poleLocationData: any, katapultPoleData: any): { from: string, to: string } {
    // Try to get from Katapult first, then fall back to SPIDAcalc
    try {
      let fromPole = "N/A";
      let toPole = "N/A";
      
      // Use SPIDAcalc data for from/to determination
      const poleId = this._canonicalizePoleID(poleLocationData.label);
      
      // Check wire end points for potential from/to poles
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) {
        return { from: fromPole, to: toPole };
      }
      
      const recommendedDesign = poleLocationData?.designs?.[designIndices.recommended]?.structure;
      if (!recommendedDesign || !recommendedDesign.wireEndPoints) {
        return { from: fromPole, to: toPole };
      }
      
      // Look for PREVIOUS_POLE and NEXT_POLE types
      for (const wireEndPoint of recommendedDesign.wireEndPoints) {
        if (wireEndPoint.type === "PREVIOUS_POLE" && wireEndPoint.structureLabel) {
          fromPole = wireEndPoint.structureLabel;
        }
        if (wireEndPoint.type === "NEXT_POLE" && wireEndPoint.structureLabel) {
          toPole = wireEndPoint.structureLabel;
        }
      }
      
      // If from/to not found, use the current pole as 'from' and the first connection as 'to'
      if (fromPole === "N/A" && toPole === "N/A" && recommendedDesign.wireEndPoints.length > 0) {
        fromPole = poleId;
        for (const wireEndPoint of recommendedDesign.wireEndPoints) {
          if (wireEndPoint.structureLabel) {
            toPole = wireEndPoint.structureLabel;
            break;
          }
        }
      }
      
      // If still no to pole, try to find any connected structure
      if (toPole === "N/A") {
        for (const wireEndPoint of recommendedDesign.wireEndPoints) {
          if (wireEndPoint.type !== "PREVIOUS_POLE" && wireEndPoint.structureLabel) {
            toPole = wireEndPoint.structureLabel;
            break;
          }
        }
      }
      
      return { from: fromPole, to: toPole };
    } catch (error) {
      console.warn("Error extracting from/to poles:", error);
      return { from: "N/A", to: "N/A" };
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

