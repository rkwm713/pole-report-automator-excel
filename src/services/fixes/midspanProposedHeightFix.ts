
/**
 * Fix for the _getMidSpanProposedHeight method to correctly calculate Column O values
 */
import { scoreWireMatch } from "../../utils/heightCalculator";

/**
 * Fix for the _getMidSpanProposedHeight method to correctly calculate Column O values
 * @param poleDataProcessor - The instance of PoleDataProcessor to patch
 */
export function fixMidspanProposedHeight(poleDataProcessor: any) {
  console.log("Applying fix for _getMidSpanProposedHeight to correct Column O values");
  
  poleDataProcessor._getMidSpanProposedHeight = function(attachment: any, pole: any) {
    if (!attachment || !pole) {
      console.log("Missing attachment or pole data for midspan height calculation");
      return "N/A";
    }
    
    try {
      // Special case for underground spans
      if (this._isUndergroundSpan(pole.id, attachment)) {
        return "UG";
      }
      
      // Get recommended design wires from SPIDAcalc
      const recommendedWires = this._getRecommendedDesignWires(pole);
      if (!recommendedWires || recommendedWires.length === 0) {
        console.log("No recommended wires found for midspan height calculation");
        return "N/A";
      }
      
      // Enhanced wire matching logic
      let bestMatch = null;
      let bestScore = -1;
      
      // Convert attachment to comparable format
      const attachmentWire = {
        owner: attachment.owner || attachment.attacherName,
        type: attachment.type || attachment.description,
        height: attachment.attachmentHeight ? 
                this._convertToInches(attachment.attachmentHeight) : undefined
      };
      
      for (const wire of recommendedWires) {
        if (!wire.midspanHeight || !wire.midspanHeight.value) continue;
        
        const wireData = {
          owner: wire.owner ? wire.owner.id : undefined,
          type: wire.clientItem ? wire.clientItem.description : undefined,
          height: wire.attachmentHeight ? 
                  this._convertToInches(wire.attachmentHeight) : undefined
        };
        
        const score = scoreWireMatch(attachmentWire, wireData);
        if (score > bestScore) {
          bestScore = score;
          bestMatch = wire;
        }
      }
      
      // If we found a good match (score threshold)
      if (bestMatch && bestScore >= 50) {
        console.log(`Found wire match with score ${bestScore} for ${attachment.description}`);
        return this._formatHeight(bestMatch.midspanHeight.value);
      }
      
      console.log(`No suitable wire match found for ${attachment.description}, best score: ${bestScore}`);
      return "N/A";
    } catch (error) {
      console.error("Error calculating midspan proposed height:", error);
      return "N/A";
    }
  };
  
  // Helper method to check if a span is underground
  poleDataProcessor._isUndergroundSpan = function(poleId: string, attachment: any): boolean {
    if (!this.katapultData || !this.katapultData.connections) {
      return false;
    }
    
    // Look for underground paths connected to this pole
    const connections = Object.values(this.katapultData.connections);
    const ugConnections = connections.filter((conn: any) => {
      if (conn && typeof conn === 'object') {
        return (
          'node_id_1' in conn && 
          'node_id_2' in conn && 
          'button' in conn &&
          (conn.node_id_1 === poleId || conn.node_id_2 === poleId) && 
          conn.button === "underground_path"
        );
      }
      return false;
    });
    
    if (ugConnections.length === 0) {
      return false;
    }
    
    // If we have attachment trace info, check if this attachment is on the underground path
    if (attachment.traceId) {
      for (const conn of ugConnections) {
        // Type guard to check if conn has sections
        if (conn && typeof conn === 'object' && 'sections' in conn && conn.sections) {
          const sections = conn.sections;
          if (typeof sections === 'object') {
            // Iterate through sections with proper type checking
            for (const sectionId in sections) {
              // Fix: Add proper type guard before accessing properties
              const section = sections[sectionId];
              if (section && typeof section === 'object') {
                // Fix: Check if attachments_on_section exists before accessing it
                if ('attachments_on_section' in section && 
                    section.attachments_on_section && 
                    typeof section.attachments_on_section === 'object') {
                  // Check if the trace ID exists in attachments_on_section
                  const attachmentsOnSection = section.attachments_on_section as Record<string, any>;
                  if (attachment.traceId in attachmentsOnSection) {
                    return true;
                  }
                }
              }
            }
          }
        }
      }
      
      // Couldn't specifically match the attachment to the UG connection
      return false;
    }
    
    // No trace ID to check, so we can't be sure - err on the side of not marking as UG
    return false;
  };
  
  // Helper method to convert height to inches
  poleDataProcessor._convertToInches = function(heightObj: any): number | undefined {
    if (!heightObj || !heightObj.value) return undefined;
    
    if (heightObj.unit === "METRE") {
      return heightObj.value * 39.3701;
    } else if (heightObj.unit === "FOOT") {
      return heightObj.value * 12;
    }
    
    return heightObj.value;
  };
}
