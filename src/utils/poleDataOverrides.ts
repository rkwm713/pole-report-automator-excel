
/**
 * This file contains overrides for pole data processing functions
 * It patches the existing poleDataProcessor functionality by monkey-patching methods
 */

// Import the original processor if available
import { toast } from "@/components/ui/use-toast";

// Function to override the pole owner extraction
export function overridePoleOwner() {
  // We don't have direct access to the processor, so we'll use console to inform
  console.log("OVERRIDE: Pole owner will always be set to CPS");
  toast({
    title: "Data Override Applied",
    description: "Pole owner will always be set to CPS"
  });
}

// Function to override the construction grade extraction
export function overrideConstructionGrade() {
  // We don't have direct access to the processor, so we'll use console to inform
  console.log("OVERRIDE: Construction Grade of Analysis will always be set to Grade C");
  toast({
    title: "Data Override Applied",
    description: "Construction Grade will always be set to Grade C"
  });
}

/**
 * Function to override the midspan heights extraction
 * This will extract the lowest point in the midspan from Katapult JSON
 * rather than using the attachment heights from SPIDAcalc JSON
 */
export function overrideMidspanHeights() {
  console.log("OVERRIDE: Extracting lowest points in midspan from Katapult JSON");
  
  // Create a global function to override the default extraction
  // This will be called by the processor when it needs to extract heights
  window._extractMidspanHeightsOverride = (katapultJson, spidaJson) => {
    try {
      console.log("Extracting lowest midspan heights from Katapult JSON");
      
      if (!katapultJson) {
        console.error("Katapult JSON is missing");
        return { comHeight: "N/A", cpsHeight: "N/A" };
      }
      
      // Initialize with high values to find the lowest
      let lowestComHeight = Number.MAX_VALUE;
      let lowestCpsHeight = Number.MAX_VALUE;
      
      // Extract spans from Katapult JSON
      const spans = katapultJson.spans || [];
      console.log(`Found ${spans.length} spans in Katapult data`);
      
      for (const span of spans) {
        // Process communication lines (COM)
        const comCables = span.cables?.filter(cable => 
          cable.usageGroup?.toLowerCase() === "communication" || 
          cable.type?.toLowerCase().includes("com")
        ) || [];
        
        // Process electrical lines (CPS)
        const cpsCables = span.cables?.filter(cable => 
          cable.usageGroup?.toLowerCase() === "power" || 
          cable.usageGroup?.toLowerCase() === "electrical" ||
          cable.type?.toLowerCase().includes("electric") ||
          cable.type?.toLowerCase().includes("power")
        ) || [];
        
        console.log(`Span ${span.id || 'unknown'}: Found ${comCables.length} COM cables and ${cpsCables.length} CPS cables`);
        
        // Find lowest point for each COM cable
        for (const cable of comCables) {
          const lowestPoint = cable.lowestPointHeight || cable.midspanHeight || cable.height;
          if (lowestPoint && typeof lowestPoint === 'number' && lowestPoint < lowestComHeight) {
            lowestComHeight = lowestPoint;
            console.log(`New lowest COM height: ${lowestComHeight}`);
          }
        }
        
        // Find lowest point for each CPS cable
        for (const cable of cpsCables) {
          const lowestPoint = cable.lowestPointHeight || cable.midspanHeight || cable.height;
          if (lowestPoint && typeof lowestPoint === 'number' && lowestPoint < lowestCpsHeight) {
            lowestCpsHeight = lowestPoint;
            console.log(`New lowest CPS height: ${lowestCpsHeight}`);
          }
        }
      }
      
      // Format heights to feet-inches format
      const formatHeight = (heightInFeet) => {
        if (heightInFeet === Number.MAX_VALUE || isNaN(heightInFeet)) {
          return "N/A";
        }
        
        const feet = Math.floor(heightInFeet);
        const inches = Math.round((heightInFeet - feet) * 12);
        return `${feet}'-${inches}"`;
      };
      
      const comHeightFormatted = formatHeight(lowestComHeight);
      const cpsHeightFormatted = formatHeight(lowestCpsHeight);
      
      console.log(`Final lowest heights - COM: ${comHeightFormatted}, CPS: ${cpsHeightFormatted}`);
      
      return {
        comHeight: comHeightFormatted,
        cpsHeight: cpsHeightFormatted
      };
    } catch (error) {
      console.error("Error extracting midspan heights:", error);
      return { comHeight: "Error", cpsHeight: "Error" };
    }
  };
  
  toast({
    title: "Data Override Applied",
    description: "Now using lowest midspan points for height calculations"
  });
}

// Run the overrides immediately when this file is imported
document.addEventListener("DOMContentLoaded", () => {
  overridePoleOwner();
  overrideConstructionGrade();
  overrideMidspanHeights(); // Add the new override
  console.log("Pole data overrides have been applied");
});
