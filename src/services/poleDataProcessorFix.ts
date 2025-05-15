
/**
 * Enhanced fixes for the PoleDataProcessor class
 * This file contains monkey patching functions that fix specific issues with the Excel output
 */

import { formatHeightString, findLowestCategorizedHeight, scoreWireMatch } from "../utils/heightCalculator";
import { WireCategory } from "../utils/katapultDataProcessor";

/**
 * Fix for the _calculateEndRow method to correctly account for From/To Pole rows
 * @param poleDataProcessor - The instance of PoleDataProcessor to patch
 */
export function fixCalculateEndRow(poleDataProcessor: any) {
  console.log("Applying fix for _calculateEndRow to correctly include From/To Pole rows");
  
  // Store the original method for calling within our patched method
  const originalCalculateEndRow = poleDataProcessor._calculateEndRow;
  
  // Replace with enhanced method
  poleDataProcessor._calculateEndRow = function(startRow: number, pole: any) {
    // Call original method to get its calculation
    const baseEndRow = originalCalculateEndRow.call(this, startRow, pole);
    
    // Add 2 more rows for From/To Pole placement at the end of each pole section
    return baseEndRow + 2;
  };
}

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

/**
 * Fix for the _writeAttachmentData method to properly place From/To Pole rows
 * @param poleDataProcessor - The instance of PoleDataProcessor to patch
 */
export function fixFromToPoleRows(poleDataProcessor: any) {
  console.log("Applying fix for _writeAttachmentData to place From/To Pole rows correctly");
  
  const originalWriteAttachmentData = poleDataProcessor._writeAttachmentData;
  
  poleDataProcessor._writeAttachmentData = function(worksheet: any, startRow: number, pole: any) {
    // Call the original method to handle regular attachments
    const lastUsedRow = originalWriteAttachmentData.call(this, worksheet, startRow, pole);
    
    // Calculate where the From/To Pole rows should go (at the end of each pole section)
    const fromPoleRow = lastUsedRow + 1;
    const toPoleRow = fromPoleRow + 1;
    
    console.log(`Writing From/To Pole rows at ${fromPoleRow} and ${toPoleRow} for pole ${pole.id || pole.poleNumber}`);
    
    // Write "From Pole" row
    worksheet.getCell(`L${fromPoleRow}`).value = "From Pole";
    worksheet.getCell(`M${fromPoleRow}`).value = pole.poleNumber || pole.id;
    worksheet.getCell(`N${fromPoleRow}`).value = "";
    worksheet.getCell(`O${fromPoleRow}`).value = "";
    
    // Write "To Pole" row
    worksheet.getCell(`L${toPoleRow}`).value = "To Pole";
    
    // Try to get the connected pole number from Katapult data
    const connectedPoleNumber = this._findConnectedPole(pole.id);
    worksheet.getCell(`M${toPoleRow}`).value = connectedPoleNumber || "N/A";
    worksheet.getCell(`N${toPoleRow}`).value = "";
    worksheet.getCell(`O${toPoleRow}`).value = "";
    
    // Return the last row used (which is now the To Pole row)
    return toPoleRow;
  };
  
  // Helper method to find connected pole
  poleDataProcessor._findConnectedPole = function(poleId: string): string | null {
    if (!this.katapultData || !this.katapultData.connections) {
      return null;
    }
    
    // Find a connection where this pole is involved
    const connections = Object.values(this.katapultData.connections);
    const connection = connections.find((conn: any) => 
      conn.node_id_1 === poleId || conn.node_id_2 === poleId
    );
    
    if (!connection) {
      return null;
    }
    
    // Get the id of the connected pole (the other end)
    const connectedPoleId = connection.node_id_1 === poleId ? 
      connection.node_id_2 : connection.node_id_1;
    
    // Look up the pole number for this connected pole
    if (connectedPoleId && this.katapultData.nodes && this.katapultData.nodes[connectedPoleId]) {
      const connectedPoleNode = this.katapultData.nodes[connectedPoleId];
      
      // Try different possible locations for pole number in Katapult data
      return (connectedPoleNode.attributes && connectedPoleNode.attributes.PoleNumber && 
              connectedPoleNode.attributes.PoleNumber.assessment) || 
             (connectedPoleNode.attributes && connectedPoleNode.attributes.pole_tag && 
              connectedPoleNode.attributes.pole_tag.tagtext) || 
             connectedPoleId;
    }
    
    return null;
  };
}

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
    const ugConnections = connections.filter((conn: any) => 
      (conn.node_id_1 === poleId || conn.node_id_2 === poleId) && 
      conn.button === "underground_path"
    );
    
    if (ugConnections.length === 0) {
      return false;
    }
    
    // If we have attachment trace info, check if this attachment is on the underground path
    if (attachment.traceId) {
      for (const conn of ugConnections) {
        if (conn.sections) {
          for (const sectionId in conn.sections) {
            const section = conn.sections[sectionId];
            if (section.attachments_on_section && 
                section.attachments_on_section[attachment.traceId]) {
              return true;
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

/**
 * Apply all fixes to the PoleDataProcessor instance
 * @param poleDataProcessor - The instance to patch
 */
export function applyAllFixes(poleDataProcessor: any) {
  console.log("Applying all fixes to PoleDataProcessor");
  fixCalculateEndRow(poleDataProcessor);
  fixExistingMidspanExtraction(poleDataProcessor);
  fixFromToPoleRows(poleDataProcessor);
  fixMidspanProposedHeight(poleDataProcessor);
  console.log("All fixes applied successfully");
}
