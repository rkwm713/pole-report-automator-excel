
/**
 * Service for processing SPIDA and Katapult data into Excel reports
 */
import XLSX from 'xlsx';

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
   * This is a placeholder implementation - the real implementation would follow the rules
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

      // This would be where we implement the full algorithm based on the rules
      // For now, we'll set up a simulated processing result
      this._simulateDataProcessing();
      
      console.log(`Processed ${this.processedPoles.length} poles`);
      return true;
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
      // Create worksheet
      const ws = XLSX.utils.aoa_to_sheet([]);
      
      // Add headers
      XLSX.utils.sheet_add_aoa(ws, [
        ["Make Ready Report"],
        [
          "Operation Number", 
          "Attachment Action:\n( I )nstalling\n( R )emoving\n( E )xisting", 
          "Pole Owner", 
          "Pole #", 
          "Pole Structure", 
          "Proposed Riser (Yes/No) &", 
          "Proposed Guy (Yes/No) &", 
          "PLA (%) with proposed attachment", 
          "Construction Grade of Analysis", 
          "Existing Mid-Span Data - Height Lowest Com", 
          "Existing Mid-Span Data - Height Lowest CPS Electrical",
          "Make Ready Data", 
          "", 
          "", 
          ""
        ],
        [
          "", "", "", "", "", "", "", "", "", "", "",
          "Attacher Description", 
          "Attachment Height - Existing", 
          "Attachment Height - Proposed", 
          "Mid-Span (same span as existing) - Proposed"
        ]
      ], { origin: "A1" });
      
      // Starting row for data
      let rowIndex = 4;
      
      // Add data for each pole
      this.processedPoles.forEach(pole => {
        const firstRowOfPole = rowIndex;
        
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
          pole.heightLowestCpsElectrical,
          "", "", "", ""
        ]], { origin: `A${rowIndex}` });
        
        rowIndex++;
        
        // Add spans and attachments
        pole.spans.forEach(span => {
          // Add span header
          XLSX.utils.sheet_add_aoa(ws, [[
            "", "", "", "", "", "", "", "", "", "", "",
            span.spanHeader, "", "", ""
          ]], { origin: `A${rowIndex}` });
          
          rowIndex++;
          
          // Add attachments
          span.attachments.forEach(attachment => {
            XLSX.utils.sheet_add_aoa(ws, [[
              "", "", "", "", "", "", "", "", "", "", "",
              attachment.description,
              attachment.existingHeight,
              attachment.proposedHeight,
              attachment.midSpanProposed
            ]], { origin: `A${rowIndex}` });
            
            rowIndex++;
          });
        });
        
        // Add From/To Pole information
        XLSX.utils.sheet_add_aoa(ws, [[
          "", "", "", "", "", "", "", "", "", "", "",
          "From Pole", pole.fromPole, "", ""
        ]], { origin: `A${rowIndex}` });
        
        rowIndex++;
        
        XLSX.utils.sheet_add_aoa(ws, [[
          "", "", "", "", "", "", "", "", "", "", "",
          "To Pole", pole.toPole, "", ""
        ]], { origin: `A${rowIndex}` });
        
        rowIndex++;
        
        // Merge cells for pole data (A-K)
        const lastRowOfPole = rowIndex - 3; // Before From Pole row
        for (let col = 0; col <= 10; col++) {
          const cellAddress = { s: { r: firstRowOfPole - 1, c: col }, e: { r: lastRowOfPole - 1, c: col } };
          if (!ws['!merges']) ws['!merges'] = [];
          ws['!merges'].push(cellAddress);
        }
      });
      
      // Set column widths
      const colWidths = [15, 20, 15, 15, 20, 15, 15, 15, 20, 20, 20, 25, 20, 20, 20];
      if (!ws['!cols']) ws['!cols'] = [];
      colWidths.forEach((width, i) => {
        ws['!cols'][i] = { wch: width };
      });
      
      // Create workbook
      const wb = XLSX.utils.book_new();
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
   * PRIVATE: Simulate data processing for demonstration
   */
  private _simulateDataProcessing(): void {
    // This simulates what the real algorithm would produce
    // In a real implementation, this would be replaced by the actual processing logic
    
    // Create some sample pole data
    const samplePoles: PoleData[] = [
      {
        operationNumber: 1,
        attachmentAction: "( I )nstalling",
        poleOwner: "Utility Co",
        poleNumber: "P1001",
        poleStructure: "45'-Class 4 Southern Pine",
        proposedRiser: "YES (1)",
        proposedGuy: "YES (2)",
        pla: "85.25%",
        constructionGrade: "Grade C",
        heightLowestCom: "18'-6\"",
        heightLowestCpsElectrical: "21'-3\"",
        spans: [
          {
            spanHeader: "Ref (North) to P1002",
            attachments: [
              {
                description: "Fiber Cable",
                existingHeight: "20'-0\"",
                proposedHeight: "22'-6\"",
                midSpanProposed: "19'-0\""
              },
              {
                description: "Power Line",
                existingHeight: "30'-0\"",
                proposedHeight: "",
                midSpanProposed: "28'-0\""
              }
            ]
          },
          {
            spanHeader: "Backspan",
            attachments: [
              {
                description: "Phone Line",
                existingHeight: "18'-0\"",
                proposedHeight: "18'-6\"",
                midSpanProposed: "17'-0\""
              }
            ]
          }
        ],
        fromPole: "P1001",
        toPole: "P1002"
      },
      {
        operationNumber: 2,
        attachmentAction: "( E )xisting",
        poleOwner: "Utility Co",
        poleNumber: "P1002",
        poleStructure: "40'-Class 5 Western Cedar",
        proposedRiser: "NO",
        proposedGuy: "YES (1)",
        pla: "67.80%",
        constructionGrade: "Grade C",
        heightLowestCom: "17'-9\"",
        heightLowestCpsElectrical: "20'-6\"",
        spans: [
          {
            spanHeader: "Ref (South) to P1001",
            attachments: [
              {
                description: "Fiber Cable",
                existingHeight: "19'-6\"",
                proposedHeight: "21'-0\"",
                midSpanProposed: "18'-6\""
              }
            ]
          },
          {
            spanHeader: "Ref (East) to P1003",
            attachments: [
              {
                description: "Power Line",
                existingHeight: "29'-0\"",
                proposedHeight: "",
                midSpanProposed: "28'-0\""
              },
              {
                description: "Communication Cable",
                existingHeight: "21'-0\"",
                proposedHeight: "22'-6\"",
                midSpanProposed: "20'-0\""
              }
            ]
          }
        ],
        fromPole: "P1002",
        toPole: "P1003"
      }
    ];

    // Assign the sample poles
    this.processedPoles = samplePoles;
  }
}
