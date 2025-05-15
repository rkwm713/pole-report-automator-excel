
/**
 * Service for processing SPIDA and Katapult data into Excel reports
 */
import { ProcessingError } from './types/poleTypes';
import { PoleDataLoader } from './dataLoader/poleDataLoader';
import { PoleDataExtractor } from './dataExtractor/poleDataExtractor';
import { PoleExcelGenerator } from './excelGenerator/poleExcelGenerator';

// Re-export the types that are used by other components
export type { PoleData, SpanData, AttachmentData, ProcessingError } from './types/poleTypes';

/**
 * Main class for processing pole data
 */
export class PoleDataProcessor {
  private dataLoader: PoleDataLoader;
  private dataExtractor: PoleDataExtractor | null = null;
  private excelGenerator: PoleExcelGenerator;
  private processedPoles: import('./types/poleTypes').PoleData[] = [];
  private errors: ProcessingError[] = [];
  private operationCounter: number = 1;

  constructor() {
    this.dataLoader = new PoleDataLoader();
    this.excelGenerator = new PoleExcelGenerator();
  }

  /**
   * Load and parse the SPIDA JSON data
   */
  loadSpidaData(jsonText: string): boolean {
    const result = this.dataLoader.loadSpidaData(jsonText);
    if (!result) {
      this.errors = this.errors.concat(this.dataLoader.getErrors());
    }
    return result;
  }

  /**
   * Load and parse the Katapult JSON data
   */
  loadKatapultData(jsonText: string): boolean {
    const result = this.dataLoader.loadKatapultData(jsonText);
    if (!result) {
      this.errors = this.errors.concat(this.dataLoader.getErrors());
    }
    return result;
  }

  /**
   * Check if all required data has been loaded
   */
  isDataLoaded(): boolean {
    return this.dataLoader.isDataLoaded();
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
      this.operationCounter = 1;

      // Create lookup maps and extract data
      const { poleLookupMap, katapultPoleLookupMap, katapultMidspanMap } = this.dataLoader.createLookupMaps();
      
      // Initialize data extractor
      this.dataExtractor = new PoleDataExtractor(poleLookupMap, katapultPoleLookupMap, katapultMidspanMap);
      
      // Get SPIDA data
      const spidaData = this.dataLoader.getSpidaData();
      
      // Process each pole from SPIDA data
      if (spidaData?.leads?.[0]?.locations) {
        console.log(`DEBUG: Found ${spidaData.leads[0].locations.length} locations/poles in SPIDA data`);
        
        for (const poleLocationData of spidaData.leads[0].locations) {
          try {
            const canonicalPoleId = this._canonicalizePoleID(poleLocationData.label);
            console.log(`DEBUG: Processing pole ${canonicalPoleId} (original label: ${poleLocationData.label})`);
            
            const katapultPoleData = katapultPoleLookupMap.get(canonicalPoleId);
            
            if (!katapultPoleData) {
              console.warn(`No matching Katapult data found for pole ${canonicalPoleId}`);
            } else {
              console.log(`DEBUG: Found matching Katapult data for pole ${canonicalPoleId}`);
            }
            
            const poleData = this.dataExtractor.extractPoleData(poleLocationData, katapultPoleData, this.operationCounter++);
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
      return this.excelGenerator.generateExcel(this.processedPoles);
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
   * PRIVATE: Helper function to canonicalize pole IDs
   */
  private _canonicalizePoleID(poleId: string): string {
    if (!poleId) return "Unknown";
    
    let id = String(poleId).trim();
    
    // Remove common prefixes
    const prefixes = ["POLE", "POLE-", "PL-", "PL", "P-", "P"];
    for (const prefix of prefixes) {
      if (id.toUpperCase().startsWith(prefix)) {
        id = id.substring(prefix.length);
        break;
      }
    }
    
    // Remove any remaining leading non-alphanumeric characters
    id = id.replace(/^[^a-zA-Z0-9]+/, "");
    
    // Handle empty ID
    return id || "Unknown";
  }
}
