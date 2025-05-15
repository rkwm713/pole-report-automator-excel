/**
 * Utility functions for processing Katapult JSON data to extract midspan heights
 * and pole attachment information, with special handling for REF connections.
 */

// Types for our data structures
interface WireObservation {
  originalHeightInches: number;
  proposedHeightInches: number | null;
}

interface WireData {
  traceId: string;
  company: string;
  cableType: string;
  midspanObservations: WireObservation[];
  poleAttachmentObservations: WireObservation[];
  lowestExistingMidspanHeight: number | null;
  finalProposedMidspanHeight: number | null;
  lowestExistingPoleAttachmentHeight: number | null;
  finalProposedPoleAttachmentHeight: number | null;
}

interface ConnectionData {
  connectionId: string;
  fromPoleId: string;
  toPoleId: string;
  buttonType: string;
  isRefConnection: boolean;
  wires: Record<string, WireData>;
}

interface ProcessedResults {
  connections: ConnectionData[];
}

/**
 * Helper function to parse height values from various formats to inches
 * @param heightValue - Height value in various formats (e.g., "22:5", "20'-7"", "240")
 * @returns Height in total inches or null if parsing fails
 */
export function parseHeightToInches(heightValue: any): number | null {
  if (heightValue === undefined || heightValue === null) {
    return null;
  }

  // If it's already a number, return it
  if (typeof heightValue === 'number') {
    return heightValue;
  }

  // Convert to string for processing
  const heightStr = String(heightValue).trim();
  if (!heightStr) {
    return null;
  }

  // Try to parse as simple number
  const numericValue = parseFloat(heightStr);
  if (!isNaN(numericValue)) {
    return numericValue;
  }

  // Parse "ft:in" format (e.g., "22:5")
  if (heightStr.includes(':')) {
    const parts = heightStr.split(':');
    if (parts.length === 2) {
      const feet = parseFloat(parts[0]);
      const inches = parseFloat(parts[1]);
      if (!isNaN(feet) && !isNaN(inches)) {
        return feet * 12 + inches;
      }
    }
  }

  // Parse "ft'-in"" format (e.g., "20'-7"")
  const feetInchesRegex = /(\d+)'[-\s]*(\d+)"/;
  const match = heightStr.match(feetInchesRegex);
  if (match) {
    const feet = parseInt(match[1], 10);
    const inches = parseInt(match[2], 10);
    return feet * 12 + inches;
  }

  // If all parsing attempts fail
  console.warn(`Failed to parse height value: ${heightStr}`);
  return null;
}

/**
 * Helper function to format height in inches to "ft'-in"" format
 * @param heightInches - Height in total inches
 * @returns Formatted height string or "N/A" if height is null/invalid
 */
export function formatHeightToString(heightInches: number | null): string {
  if (heightInches === null || isNaN(heightInches)) {
    return "N/A";
  }

  const feet = Math.floor(heightInches / 12);
  const inches = Math.round(heightInches % 12);
  
  // Handle case where inches rounds up to 12
  if (inches === 12) {
    return `${feet + 1}'-0"`;
  }
  
  return `${feet}'-${inches}"`;
}

/**
 * Process Katapult JSON data to extract and analyze midspan heights and pole attachment information
 * @param katapultJson - The parsed Katapult JSON data
 * @returns Processed results with connection and wire data
 */
export function processKatapultData(katapultJson: any): ProcessedResults {
  console.log("Starting Katapult data processing");
  
  // Initialize result structure
  const results: ProcessedResults = {
    connections: []
  };
  
  if (!katapultJson || typeof katapultJson !== 'object') {
    console.error("Invalid Katapult JSON data");
    return results;
  }

  // Access required data structures
  const connections = katapultJson.connections || {};
  const photoSummary = katapultJson.photo_summary || {};
  const traces = katapultJson.traces?.trace_data || {};
  const nodes = katapultJson.nodes || {};

  console.log(`Found ${Object.keys(connections).length} connections to process`);
  
  // Process each connection
  for (const connectionId in connections) {
    const connectionData = connections[connectionId];
    
    // Skip if connection data is invalid
    if (!connectionData || typeof connectionData !== 'object') {
      continue;
    }
    
    // Extract connection metadata
    const fromPoleId = connectionData.node_id_1 || '';
    const toPoleId = connectionData.node_id_2 || '';
    const buttonType = connectionData.button_type || '';
    const isRefConnection = buttonType === 'ref';
    
    console.log(`Processing connection ${connectionId}: ${fromPoleId} to ${toPoleId} (${buttonType})`);
    
    // Initialize connection structure
    const connection: ConnectionData = {
      connectionId,
      fromPoleId,
      toPoleId,
      buttonType,
      isRefConnection,
      wires: {}
    };
    
    // Process midspan sections for this connection
    if (connectionData.sections && typeof connectionData.sections === 'object') {
      for (const sectionId in connectionData.sections) {
        const sectionData = connectionData.sections[sectionId];
        
        // Skip if section photos are invalid
        if (!sectionData.photos || typeof sectionData.photos !== 'object') {
          continue;
        }
        
        // Process each photo in the section
        for (const photoId in sectionData.photos) {
          // Get photo details from the photo summary
          const photoDetail = photoSummary[photoId];
          if (!photoDetail || !photoDetail.photofirst_data) {
            continue;
          }
          
          const photofirstData = photoDetail.photofirst_data;
          
          // Process wires in the photo
          if (photofirstData.wire && typeof photofirstData.wire === 'object') {
            for (const wireInstanceId in photofirstData.wire) {
              const wireInstanceData = photofirstData.wire[wireInstanceId];
              
              // Get trace ID (skip if missing)
              const traceId = wireInstanceData.trace_id;
              if (!traceId) {
                continue;
              }
              
              // Initialize wire data if not exists
              if (!connection.wires[traceId]) {
                // Get wire metadata from traces
                const traceData = traces[traceId] || {};
                
                connection.wires[traceId] = {
                  traceId,
                  company: traceData.company || 'Unknown',
                  cableType: traceData.cable_type || 'Unknown',
                  midspanObservations: [],
                  poleAttachmentObservations: [],
                  lowestExistingMidspanHeight: null,
                  finalProposedMidspanHeight: null,
                  lowestExistingPoleAttachmentHeight: null,
                  finalProposedPoleAttachmentHeight: null
                };
              }
              
              // Parse wire height (prioritize _manual_height)
              const heightValue = wireInstanceData._manual_height !== undefined ? 
                wireInstanceData._manual_height : 
                wireInstanceData._measured_height;
              
              const currentHeightInches = parseHeightToInches(heightValue);
              
              // Process height and proposed moves
              if (currentHeightInches !== null) {
                // Check for proposed moves
                const mrMoveVal = wireInstanceData.mr_move;
                const effectiveMovesVal = wireInstanceData._effective_moves;
                
                let isInstanceProposed = false;
                let calculatedProposedHeightForInstance = currentHeightInches;
                
                // Apply mr_move if present and not zero
                if (mrMoveVal !== undefined && mrMoveVal !== null && 
                    !isNaN(parseFloat(mrMoveVal)) && parseFloat(mrMoveVal) !== 0) {
                  calculatedProposedHeightForInstance = currentHeightInches + parseFloat(mrMoveVal);
                  isInstanceProposed = true;
                }
                // Otherwise check for effective moves
                else if (effectiveMovesVal !== undefined && effectiveMovesVal !== null && 
                        effectiveMovesVal !== '') {
                  isInstanceProposed = true;
                  // Note: height remains the same as this is assumed to already reflect the moves
                }
                
                // Store observation in midspan observations
                connection.wires[traceId].midspanObservations.push({
                  originalHeightInches: currentHeightInches,
                  proposedHeightInches: isInstanceProposed ? calculatedProposedHeightForInstance : null
                });
              }
            }
          }
        }
      }
    }
    
    // Process pole attachments for REF connections
    if (isRefConnection && fromPoleId) {
      const poleData = nodes[fromPoleId];
      
      if (poleData && poleData.photos && typeof poleData.photos === 'object') {
        for (const photoId in poleData.photos) {
          // Get photo details
          const photoDetail = photoSummary[photoId];
          if (!photoDetail || !photoDetail.photofirst_data) {
            continue;
          }
          
          const photofirstData = photoDetail.photofirst_data;
          
          // Process wires in the photo
          if (photofirstData.wire && typeof photofirstData.wire === 'object') {
            for (const wireInstanceId in photofirstData.wire) {
              const wireInstanceData = photofirstData.wire[wireInstanceId];
              
              // Get trace ID (skip if missing)
              const traceId = wireInstanceData.trace_id;
              if (!traceId) {
                continue;
              }
              
              // Initialize wire data if not exists
              if (!connection.wires[traceId]) {
                // Get wire metadata from traces
                const traceData = traces[traceId] || {};
                
                connection.wires[traceId] = {
                  traceId,
                  company: traceData.company || 'Unknown',
                  cableType: traceData.cable_type || 'Unknown',
                  midspanObservations: [],
                  poleAttachmentObservations: [],
                  lowestExistingMidspanHeight: null,
                  finalProposedMidspanHeight: null,
                  lowestExistingPoleAttachmentHeight: null,
                  finalProposedPoleAttachmentHeight: null
                };
              }
              
              // Parse wire attachment height
              const heightValue = wireInstanceData._manual_height !== undefined ? 
                wireInstanceData._manual_height : 
                wireInstanceData._measured_height;
              
              const currentHeightInches = parseHeightToInches(heightValue);
              
              // Process height and proposed moves
              if (currentHeightInches !== null) {
                // Check for proposed moves
                const mrMoveVal = wireInstanceData.mr_move;
                const effectiveMovesVal = wireInstanceData._effective_moves;
                
                let isInstanceProposed = false;
                let calculatedProposedHeightForInstance = currentHeightInches;
                
                // Apply mr_move if present and not zero
                if (mrMoveVal !== undefined && mrMoveVal !== null && 
                    !isNaN(parseFloat(mrMoveVal)) && parseFloat(mrMoveVal) !== 0) {
                  calculatedProposedHeightForInstance = currentHeightInches + parseFloat(mrMoveVal);
                  isInstanceProposed = true;
                }
                // Otherwise check for effective moves
                else if (effectiveMovesVal !== undefined && effectiveMovesVal !== null && 
                        effectiveMovesVal !== '') {
                  isInstanceProposed = true;
                  // Note: height remains the same as this is assumed to already reflect the moves
                }
                
                // Store observation in pole attachment observations
                connection.wires[traceId].poleAttachmentObservations.push({
                  originalHeightInches: currentHeightInches,
                  proposedHeightInches: isInstanceProposed ? calculatedProposedHeightForInstance : null
                });
              }
            }
          }
        }
      }
    }
    
    // Aggregate and finalize wire data
    for (const traceId in connection.wires) {
      const wire = connection.wires[traceId];
      
      // Calculate lowest existing midspan height
      const existingMidspanHeights = wire.midspanObservations
        .map(obs => obs.originalHeightInches)
        .filter(height => height !== null) as number[];
      
      if (existingMidspanHeights.length > 0) {
        wire.lowestExistingMidspanHeight = Math.min(...existingMidspanHeights);
      }
      
      // Calculate final proposed midspan height
      const proposedMidspanHeights = wire.midspanObservations
        .map(obs => obs.proposedHeightInches)
        .filter(height => height !== null) as number[];
      
      if (proposedMidspanHeights.length > 0) {
        wire.finalProposedMidspanHeight = Math.min(...proposedMidspanHeights);
      }
      
      // For REF connections, calculate pole attachment heights
      if (isRefConnection) {
        // Calculate lowest existing pole attachment height
        const existingPoleAttachmentHeights = wire.poleAttachmentObservations
          .map(obs => obs.originalHeightInches)
          .filter(height => height !== null) as number[];
        
        if (existingPoleAttachmentHeights.length > 0) {
          wire.lowestExistingPoleAttachmentHeight = Math.min(...existingPoleAttachmentHeights);
        }
        
        // Calculate final proposed pole attachment height
        const proposedPoleAttachmentHeights = wire.poleAttachmentObservations
          .map(obs => obs.proposedHeightInches)
          .filter(height => height !== null) as number[];
        
        if (proposedPoleAttachmentHeights.length > 0) {
          wire.finalProposedPoleAttachmentHeight = Math.min(...proposedPoleAttachmentHeights);
        }
      }
    }
    
    // Add connection to results
    results.connections.push(connection);
  }
  
  console.log(`Processed ${results.connections.length} connections with wire data`);
  return results;
}

/**
 * Generate a formatted summary of the processed Katapult data
 * @param results - The processed results from processKatapultData
 * @returns A formatted string summary
 */
export function generateKatapultDataSummary(results: ProcessedResults): string {
  let summary = "# Katapult Data Analysis Summary\n\n";
  
  for (const connection of results.connections) {
    summary += `## Connection: ${connection.connectionId}\n`;
    summary += `- From Pole: ${connection.fromPoleId}\n`;
    summary += `- To Pole: ${connection.toPoleId}\n`;
    summary += `- Type: ${connection.buttonType}\n`;
    
    if (connection.isRefConnection) {
      summary += `- **REF Connection** (FROM POLE: ${connection.fromPoleId})\n`;
    }
    
    summary += "\n### Wire Data:\n\n";
    
    for (const traceId in connection.wires) {
      const wire = connection.wires[traceId];
      
      summary += `#### ${wire.company} - ${wire.cableType} (Trace ID: ${wire.traceId})\n`;
      
      // Midspan Heights
      summary += "- Mid-Span Heights:\n";
      const existingMidspanHeight = wire.lowestExistingMidspanHeight !== null ? 
        formatHeightToString(wire.lowestExistingMidspanHeight) : "N/A";
      const proposedMidspanHeight = wire.finalProposedMidspanHeight !== null ? 
        formatHeightToString(wire.finalProposedMidspanHeight) : "N/A";
      
      if (wire.finalProposedMidspanHeight !== null || connection.isRefConnection) {
        summary += `  - Existing: (${existingMidspanHeight})\n`;
      } else {
        summary += `  - Existing: ${existingMidspanHeight}\n`;
      }
      
      summary += `  - Proposed: ${proposedMidspanHeight}\n`;
      
      // Pole Attachment Heights (only for REF connections)
      if (connection.isRefConnection) {
        summary += "- Pole Attachment Heights:\n";
        const existingPoleHeight = wire.lowestExistingPoleAttachmentHeight !== null ? 
          formatHeightToString(wire.lowestExistingPoleAttachmentHeight) : "N/A";
        const proposedPoleHeight = wire.finalProposedPoleAttachmentHeight !== null ? 
          formatHeightToString(wire.finalProposedPoleAttachmentHeight) : "N/A";
        
        summary += `  - Existing: (${existingPoleHeight})\n`;
        summary += `  - Proposed: ${proposedPoleHeight}\n`;
      }
      
      summary += "\n";
    }
    
    summary += "---\n\n";
  }
  
  return summary;
}
