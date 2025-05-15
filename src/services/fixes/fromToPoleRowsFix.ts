
/**
 * Fix for the _writeAttachmentData method to properly place From/To Pole rows
 */

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
    
    // Add proper type guard before accessing properties
    const connection = connections.find((conn: any) => {
      if (conn && typeof conn === 'object') {
        return (
          'node_id_1' in conn && 
          'node_id_2' in conn && 
          (conn.node_id_1 === poleId || conn.node_id_2 === poleId)
        );
      }
      return false;
    });
    
    if (!connection || typeof connection !== 'object') {
      return null;
    }
    
    // Type guard to ensure connection has required properties
    if (!('node_id_1' in connection && 'node_id_2' in connection)) {
      return null;
    }
    
    // Get the id of the connected pole (the other end)
    const connectedPoleId = connection.node_id_1 === poleId ? 
      connection.node_id_2 : connection.node_id_1;
    
    // Look up the pole number for this connected pole
    if (connectedPoleId && this.katapultData.nodes) {
      // Check if nodes has the connectedPoleId as a key
      const nodes = this.katapultData.nodes as Record<string, any>;
      if (!(connectedPoleId in nodes)) {
        return null;
      }
      
      const connectedPoleNode = nodes[connectedPoleId];
      if (!connectedPoleNode || typeof connectedPoleNode !== 'object') {
        return null;
      }
      
      // Try different possible locations for pole number in Katapult data
      if (connectedPoleNode.attributes && typeof connectedPoleNode.attributes === 'object') {
        if ('PoleNumber' in connectedPoleNode.attributes && 
            connectedPoleNode.attributes.PoleNumber && 
            typeof connectedPoleNode.attributes.PoleNumber === 'object' &&
            'assessment' in connectedPoleNode.attributes.PoleNumber) {
          return connectedPoleNode.attributes.PoleNumber.assessment as string;
        } else if ('pole_tag' in connectedPoleNode.attributes && 
                  connectedPoleNode.attributes.pole_tag && 
                  typeof connectedPoleNode.attributes.pole_tag === 'object' &&
                  'tagtext' in connectedPoleNode.attributes.pole_tag) {
          return connectedPoleNode.attributes.pole_tag.tagtext as string;
        }
      }
      
      return connectedPoleId as string;
    }
    
    return null;
  };
}
