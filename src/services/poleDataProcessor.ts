
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
