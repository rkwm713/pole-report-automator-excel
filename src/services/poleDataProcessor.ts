
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

      // Add debug logging for Katapult data structure
      console.log("DEBUG: Katapult data structure:", JSON.stringify(this.katapultData).substring(0, 500) + "...");
      
      if (typeof this.katapultData === 'object') {
        console.log("DEBUG: Katapult data keys:", Object.keys(this.katapultData));
      }
      
      // Create lookup maps for efficient matching between data sources
      this._createPoleLookupMaps();
      
      // Process each pole from SPIDA data
      if (this.spidaData?.leads?.[0]?.locations) {
        for (const poleLocationData of this.spidaData.leads[0].locations) {
          try {
            const canonicalPoleId = this._canonicalizePoleID(poleLocationData.label);
            const katapultPoleData = this.katapultPoleLookupMap.get(canonicalPoleId);
            
            if (!katapultPoleData) {
              console.warn(`No matching Katapult data found for pole ${canonicalPoleId}`);
            }
            
            const poleData = this._extractPoleData(poleLocationData, katapultPoleData);
            if (poleData) {
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
    
    // Apply styling - Make header rows bold
    if (!ws['!rows']) ws['!rows'] = [];
    for (let i = 0; i < 3; i++) {
      ws['!rows'][i] = { hidden: false, hpt: 20 }; // Set row height
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
      // Write span header
      XLSX.utils.sheet_add_aoa(ws, [[
        span.spanHeader, "", "", ""
      ]], { origin: `L${currentRow}` });
      
      currentRow++;
      
      // Write attachments
      for (const attachment of span.attachments) {
        XLSX.utils.sheet_add_aoa(ws, [[
          attachment.description,
          attachment.existingHeight,
          attachment.proposedHeight,
          attachment.midSpanProposed
        ]], { origin: `L${currentRow}` });
        
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
    
    currentRow++;
    
    // Add "To Pole" row
    XLSX.utils.sheet_add_aoa(ws, [[
      "", "", "", "", "", "", "", "", "", "", "",
      "To Pole", pole.toPole, "", ""
    ]], { origin: `A${currentRow}` });
    
    return currentRow + 1;
  }
  
  /**
   * PRIVATE: Set column widths
   */
  private _setColumnWidths(ws: XLSX.WorkSheet): void {
    if (!ws['!cols']) ws['!cols'] = [];
    
    // Set specific widths for each column
    const colWidths = [
      15, // A: Operation Number
      20, // B: Attachment Action
      15, // C: Pole Owner
      15, // D: Pole #
      25, // E: Pole Structure (wider)
      17, // F: Proposed Riser
      17, // G: Proposed Guy
      15, // H: PLA (%)
      20, // I: Construction Grade
      20, // J: Height Lowest Com
      20, // K: Height Lowest CPS Electrical
      30, // L: Attacher Description (wider)
      15, // M: Existing
      15, // N: Proposed
      20, // O: Mid-Span Proposed
    ];
    
    // Apply column widths
    colWidths.forEach((width, i) => {
      ws['!cols'][i] = { wch: width };
    });
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
    }
    
    // Create maps for Katapult poles
    // Make this method more robust to handle different Katapult data structures
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
      // Find design indices
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) {
        throw new Error("Could not find required designs in pole data");
      }
      
      // Extract basic pole information
      const canonicalPoleId = this._canonicalizePoleID(poleLocationData.label);
      
      // Extract pole owner (prioritize Katapult if available)
      let poleOwner = "Unknown";
      if (katapultPoleData?.properties?.poleOwner) {
        poleOwner = katapultPoleData.properties.poleOwner;
      } else if (poleLocationData?.structure?.pole?.owner?.id) {
        poleOwner = poleLocationData.structure.pole.owner.id;
      }
      
      // Extract pole structure details
      const poleStructure = this._extractPoleStructure(poleLocationData);
      
      // Extract proposed riser information
      const proposedRiser = this._extractProposedRiserInfo(poleLocationData, designIndices.recommended);
      
      // Extract proposed guy information
      const proposedGuy = this._extractProposedGuyInfo(poleLocationData, designIndices.recommended);
      
      // Extract PLA value
      const plaInfo = this._extractPLA(poleLocationData);
      
      // Extract construction grade
      const constructionGrade = this._extractConstructionGrade(poleLocationData);
      
      // Extract midspan height data
      const midspanHeights = this._extractExistingMidspanData(poleLocationData, katapultPoleData);
      
      // Extract span data with attachments
      const spans = this._extractSpanData(poleLocationData, designIndices);
      
      // Extract from/to pole information (prioritize Katapult if available)
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
   * PRIVATE: Find indices for "Measured Design" and "Recommended Design"
   */
  private _findDesignIndices(poleLocationData: any): { measured: number, recommended: number } | null {
    if (!poleLocationData?.designs || !Array.isArray(poleLocationData.designs)) {
      return null;
    }
    
    let measuredIdx = -1;
    let recommendedIdx = -1;
    
    for (let i = 0; i < poleLocationData.designs.length; i++) {
      const design = poleLocationData.designs[i];
      if (design?.name?.includes("Measured Design")) {
        measuredIdx = i;
      }
      if (design?.name?.includes("Recommended Design")) {
        recommendedIdx = i;
      }
    }
    
    // Fallback to first two designs if specific names not found
    if (measuredIdx === -1 && poleLocationData.designs.length > 0) {
      measuredIdx = 0;
    }
    if (recommendedIdx === -1 && poleLocationData.designs.length > 1) {
      recommendedIdx = 1;
    } else if (recommendedIdx === -1 && measuredIdx === 0) {
      // If only one design, use it for both
      recommendedIdx = 0;
    }
    
    if (measuredIdx === -1 || recommendedIdx === -1) {
      return null;
    }
    
    return { measured: measuredIdx, recommended: recommendedIdx };
  }

  /**
   * PRIVATE: Extract pole structure information
   */
  private _extractPoleStructure(poleLocationData: any): string {
    try {
      const pole = poleLocationData?.structure?.pole;
      if (!pole) return "Unknown";
      
      // Get height and convert from meters to feet
      let height = "";
      if (pole.clientItem?.height?.value) {
        const meters = pole.clientItem.height.value;
        const feet = this._metersToFeet(meters);
        height = `${Math.round(feet)}'`;
      }
      
      // Get class
      const poleClass = pole.clientItem?.classOfPole || "";
      
      // Get species
      const species = pole.clientItem?.species || "";
      
      return `${height}-Class ${poleClass} ${species}`.trim();
    } catch (error) {
      console.warn("Error extracting pole structure:", error);
      return "Unknown";
    }
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
   */
  private _extractPLA(poleLocationData: any): { pla: string, actual: number } {
    try {
      // Find the target analysis case (assuming "Light - Grade C" or similar)
      const analysisObjects = poleLocationData?.analysis;
      if (!analysisObjects || !Array.isArray(analysisObjects)) {
        return { pla: "N/A", actual: 0 };
      }
      
      // Look for the recommended design analysis case
      let targetAnalysis = null;
      for (const analysis of analysisObjects) {
        const caseName = analysis?.analysisCaseDetails?.name || "";
        if (
          caseName.includes("Recommended") || 
          caseName.includes("Light - Grade") ||
          caseName.includes("Grade C")
        ) {
          targetAnalysis = analysis;
          break;
        }
      }
      
      // Use the first analysis if specific case not found
      if (!targetAnalysis && analysisObjects.length > 0) {
        targetAnalysis = analysisObjects[0];
      }
      
      // If still no analysis found
      if (!targetAnalysis) {
        return { pla: "N/A", actual: 0 };
      }
      
      // Find the pole stress result
      const results = targetAnalysis.results || [];
      for (const result of results) {
        if (
          result.component === "Pole" && 
          result.analysisType === "STRESS"
        ) {
          const plaValue = result.actual;
          if (typeof plaValue === "number") {
            // Format as percentage with 2 decimal places
            return { 
              pla: `${plaValue.toFixed(2)}%`, 
              actual: plaValue 
            };
          }
        }
      }
      
      return { pla: "N/A", actual: 0 };
    } catch (error) {
      console.warn("Error extracting PLA:", error);
      return { pla: "N/A", actual: 0 };
    }
  }

  /**
   * PRIVATE: Extract construction grade
   */
  private _extractConstructionGrade(poleLocationData: any): string {
    try {
      // Find the target analysis case
      const analysisObjects = poleLocationData?.analysis;
      if (!analysisObjects || !Array.isArray(analysisObjects)) {
        return "N/A";
      }
      
      // Look for the recommended design analysis case
      for (const analysis of analysisObjects) {
        const constructionGrade = analysis?.analysisCaseDetails?.constructionGrade;
        if (constructionGrade) {
          return `Grade ${constructionGrade}`;
        }
      }
      
      return "N/A";
    } catch (error) {
      console.warn("Error extracting construction grade:", error);
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
   */
  private _canonicalizePoleID(poleID: string): string {
    if (!poleID) return "Unknown";
    
    // Extract core pole ID by removing prefixes, etc.
    // This implementation will need to be adjusted based on actual pole ID formats
    
    // Example: strip prefix like "1-PL410620" to get "PL410620"
    // This is just a sample implementation and should be adjusted
    
    const match = poleID.match(/[A-Z]{2}\d+/);
    if (match) {
      return match[0];
    }
    
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
