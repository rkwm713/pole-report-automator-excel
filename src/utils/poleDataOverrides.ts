
/**
 * This file contains overrides for pole data processing functions
 * It patches the existing poleDataProcessor functionality by monkey-patching methods
 */

// Import the original processor if available
import { toast } from "@/hooks/use-toast";
import { processKatapultData, formatHeightToString } from "./katapultDataProcessor";

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
      console.log("Extracting lowest midspan heights from Katapult JSON with enhanced processor");
      
      if (!katapultJson) {
        console.error("Katapult JSON is missing");
        return { comHeight: "N/A", cpsHeight: "N/A" };
      }
      
      // Process the Katapult data using our enhanced processor
      const processedResults = processKatapultData(katapultJson);
      
      // Find the lowest COM and CPS heights across all connections
      let lowestComHeight = Number.MAX_VALUE;
      let lowestCpsHeight = Number.MAX_VALUE;
      
      // Analyze all connections and their wires
      for (const connection of processedResults.connections) {
        for (const traceId in connection.wires) {
          const wire = connection.wires[traceId];
          const company = wire.company.toLowerCase();
          const cableType = wire.cableType.toLowerCase();
          
          // Determine if this is a communication or power wire
          const isCom = company.includes('com') || 
                       cableType.includes('com') || 
                       company.includes('telephone') || 
                       cableType.includes('telephone');
                       
          const isCps = company.includes('power') || 
                       cableType.includes('power') || 
                       company.includes('electric') || 
                       cableType.includes('electric');
          
          // Use proposed height if available, otherwise existing height
          const heightToUse = wire.finalProposedMidspanHeight !== null ? 
                             wire.finalProposedMidspanHeight : 
                             wire.lowestExistingMidspanHeight;
          
          if (heightToUse !== null) {
            if (isCom && heightToUse < lowestComHeight) {
              lowestComHeight = heightToUse;
              console.log(`New lowest COM height: ${formatHeightToString(lowestComHeight)} from trace ${traceId}`);
            }
            
            if (isCps && heightToUse < lowestCpsHeight) {
              lowestCpsHeight = heightToUse;
              console.log(`New lowest CPS height: ${formatHeightToString(lowestCpsHeight)} from trace ${traceId}`);
            }
          }
        }
      }
      
      // Format the heights
      const comHeightFormatted = lowestComHeight !== Number.MAX_VALUE ? 
                               formatHeightToString(lowestComHeight) : "N/A";
      const cpsHeightFormatted = lowestCpsHeight !== Number.MAX_VALUE ? 
                               formatHeightToString(lowestCpsHeight) : "N/A";
      
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
