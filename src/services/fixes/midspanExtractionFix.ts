
/**
 * Fix for the extraction of existing midspan data for Columns J & K
 */
import { formatHeightString, findLowestCategorizedHeight } from "../../utils/heightCalculator";
import { WireCategory } from "../../utils/katapultDataProcessor";

/**
 * Fix for the extraction of existing midspan data for Columns J & K
 * @param poleDataProcessor - The instance of PoleDataProcessor to patch
 */
export function fixExistingMidspanExtraction(poleDataProcessor: any) {
  console.log("Applying fix for _extractExistingMidspanData to correct height values in columns J & K");
  
  poleDataProcessor._extractExistingMidspanData = function(pole: any) {
    try {
      console.log(`Extracting existing midspan data for pole ${pole.id || pole.poleNumber}`);
      
      // Use the Katapult data if available through our override
      if (typeof window !== 'undefined' && window._extractMidspanHeightsOverride && this.katapultData) {
        console.log("Using enhanced Katapult midspan height extraction");
        const result = window._extractMidspanHeightsOverride(this.katapultData, this.spidaData);
        return result;
      }
      
      // Fallback to simplified extraction if no override available
      const connections = this._findKatapultConnectionsForPole(pole.id);
      if (!connections || connections.length === 0) {
        console.log("No connections found for pole, returning N/A for heights");
        return { comHeight: "N/A", cpsHeight: "N/A" };
      }
      
      let lowestComHeight = Number.MAX_VALUE;
      let lowestCpsHeight = Number.MAX_VALUE;
      let hasComHeights = false;
      let hasCpsHeights = false;
      
      // Process each connection to find lowest heights
      connections.forEach(connection => {
        if (!connection.wires) return;
        
        Object.values(connection.wires).forEach((wire: any) => {
          if (!wire.category || !wire.lowestExistingMidspanHeight) return;
          
          if (wire.category === WireCategory.COMMUNICATION && wire.lowestExistingMidspanHeight < lowestComHeight) {
            lowestComHeight = wire.lowestExistingMidspanHeight;
            hasComHeights = true;
            console.log(`Found lower COM height: ${formatHeightString(lowestComHeight)} from ${wire.company} ${wire.cableType}`);
          }
          
          if (wire.category === WireCategory.CPS_ELECTRICAL && wire.lowestExistingMidspanHeight < lowestCpsHeight) {
            lowestCpsHeight = wire.lowestExistingMidspanHeight;
            hasCpsHeights = true;
            console.log(`Found lower CPS height: ${formatHeightString(lowestCpsHeight)} from ${wire.company} ${wire.cableType}`);
          }
        });
      });
      
      // Format the heights or return N/A if none found
      const comHeightFormatted = hasComHeights ? formatHeightString(lowestComHeight) : "N/A";
      const cpsHeightFormatted = hasCpsHeights ? formatHeightString(lowestCpsHeight) : "N/A";
      
      console.log(`Final lowest heights - COM: ${comHeightFormatted}, CPS: ${cpsHeightFormatted}`);
      
      return {
        comHeight: comHeightFormatted,
        cpsHeight: cpsHeightFormatted
      };
    } catch (error) {
      console.error("Error extracting midspan heights:", error);
      return { comHeight: "N/A", cpsHeight: "N/A" };
    }
  };
}
