
/**
 * PoleData interface represents the processed pole data structure
 */
export interface PoleData {
  poleNumber: string;
  operationNumber: string;
  poleOwner: string;
  attachmentAction: string;
  pla: string;
  spans: Array<any>;
}

/**
 * PoleDataProcessor class contains methods to extract construction grade
 */
export class PoleDataProcessor {
  private spidaData: any = null;
  private katapultData: any = null;
  private processedPoles: PoleData[] = [];
  private errors: Array<{ code: string; message: string; details?: string }> = [];
  
  /**
   * Load SPIDA Data
   */
  loadSpidaData(jsonContent: string): boolean {
    try {
      this.spidaData = JSON.parse(jsonContent);
      return true;
    } catch (error) {
      this.errors.push({
        code: 'SPIDA_PARSE_ERROR',
        message: 'Failed to parse SPIDA data',
        details: error instanceof Error ? error.message : String(error)
      });
      return false;
    }
  }
  
  /**
   * Load Katapult Data
   */
  loadKatapultData(jsonContent: string): boolean {
    try {
      this.katapultData = JSON.parse(jsonContent);
      return true;
    } catch (error) {
      this.errors.push({
        code: 'KATAPULT_PARSE_ERROR',
        message: 'Failed to parse Katapult data',
        details: error instanceof Error ? error.message : String(error)
      });
      return false;
    }
  }
  
  /**
   * Process Data
   */
  processData(): boolean {
    try {
      // Reset any previous processing
      this.processedPoles = [];
      this.errors = [];
      
      // Check if data is loaded
      if (!this.spidaData || !this.katapultData) {
        this.errors.push({
          code: 'NO_DATA',
          message: 'Both SPIDA and Katapult data must be loaded before processing'
        });
        return false;
      }
      
      // For now, just create some sample processed poles
      this.processedPoles = [
        {
          poleNumber: "P123",
          operationNumber: "OP-456",
          poleOwner: "Utility Co.",
          attachmentAction: "Transfer",
          pla: "Yes",
          spans: []
        },
        {
          poleNumber: "P124",
          operationNumber: "OP-457",
          poleOwner: "Telecom Inc.",
          attachmentAction: "New",
          pla: "No",
          spans: []
        }
      ];
      
      return true;
    } catch (error) {
      this.errors.push({
        code: 'PROCESSING_ERROR',
        message: 'Error processing data',
        details: error instanceof Error ? error.message : String(error)
      });
      return false;
    }
  }
  
  /**
   * Get Processed Poles
   */
  getProcessedPoles(): PoleData[] {
    return this.processedPoles;
  }
  
  /**
   * Get Processed Pole Count
   */
  getProcessedPoleCount(): number {
    return this.processedPoles.length;
  }
  
  /**
   * Get Errors
   */
  getErrors(): Array<{ code: string; message: string; details?: string }> {
    return this.errors;
  }
  
  /**
   * Generate Excel
   */
  generateExcel(): Blob | null {
    if (this.processedPoles.length === 0) {
      return null;
    }
    
    // In a real implementation, this would generate an actual Excel file
    // For now, just return a simple text blob
    return new Blob(['Sample Excel Data'], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  
  /**
   * PRIVATE: Extract construction grade
   */
  private _extractConstructionGrade(poleLocationData: any): string {
    try {
      console.log("DEBUG: Extracting construction grade");
      
      // Find design indices
      const designIndices = this._findDesignIndices(poleLocationData);
      if (!designIndices) {
        return "Grade C"; // Default to Grade C if no design indices found
      }
      
      // Try to get from analysis case name first
      if (poleLocationData.analysis && Array.isArray(poleLocationData.analysis)) {
        for (const analysisCase of poleLocationData.analysis) {
          const caseName = analysisCase?.analysisCaseDetails?.name || "";
          
          if (caseName.includes("Grade")) {
            // Extract the grade from the case name
            const gradeMatch = caseName.match(/Grade\s+([A-F])/i);
            if (gradeMatch && gradeMatch[1]) {
              console.log(`DEBUG: Found construction grade ${gradeMatch[1]} in analysis case name`);
              return `Grade ${gradeMatch[1]}`;
            }
          }
        }
      }
      
      // If no grade found in analysis case names, look in the recommended design
      const recommendedDesignIdx = designIndices.recommended;
      const recommendedDesign = poleLocationData?.designs?.[recommendedDesignIdx];
      
      if (recommendedDesign?.analysis && Array.isArray(recommendedDesign.analysis)) {
        for (const analysisCase of recommendedDesign.analysis) {
          const caseName = analysisCase?.analysisCaseDetails?.name || "";
          
          if (caseName.includes("Grade")) {
            // Extract the grade from the case name
            const gradeMatch = caseName.match(/Grade\s+([A-F])/i);
            if (gradeMatch && gradeMatch[1]) {
              console.log(`DEBUG: Found construction grade ${gradeMatch[1]} in design analysis case`);
              return `Grade ${gradeMatch[1]}`;
            }
          }
        }
      }
      
      // Default to Grade C if not found
      console.log("DEBUG: No construction grade found, defaulting to Grade C");
      return "Grade C";
    } catch (error) {
      console.warn("Error extracting construction grade:", error);
      return "Grade C"; // Default to Grade C on error
    }
  }

  /**
   * PRIVATE: Find design indices in pole location data
   */
  private _findDesignIndices(poleLocationData: any): { recommended: number } | null {
    try {
      if (!poleLocationData.designs || !Array.isArray(poleLocationData.designs)) {
        return null;
      }
      
      // Find the recommended design
      let recommendedIndex = -1;
      for (let i = 0; i < poleLocationData.designs.length; i++) {
        if (poleLocationData.designs[i].status === "recommended") {
          recommendedIndex = i;
          break;
        }
      }
      
      if (recommendedIndex === -1) {
        return null;
      }
      
      return {
        recommended: recommendedIndex
      };
    } catch (error) {
      console.warn("Error finding design indices:", error);
      return null;
    }
  }
}

// Export an instance of the class for convenience
export const poleDataProcessor = new PoleDataProcessor();
