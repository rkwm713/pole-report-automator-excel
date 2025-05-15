
import * as XLSX from 'xlsx';
import { WireCategory, formatHeightToString } from '@/utils/katapultDataProcessor';

/**
 * Generates and downloads an Excel file with the analysis of Katapult data
 * @param processedData - The processed Katapult data
 * @param rawKatapultData - The raw Katapult JSON data
 */
export function downloadKatapultAnalysisExcel(processedData: any, rawKatapultData: any) {
  // Create a new workbook
  const workbook = XLSX.utils.book_new();
  
  // Create sheets
  addSummarySheet(workbook, processedData);
  addConnectionsSheet(workbook, processedData);
  addWireDataSheet(workbook, processedData);
  addRefConnectionsSheet(workbook, processedData);
  
  // Generate and download Excel file
  XLSX.writeFile(workbook, 'katapult-analysis.xlsx');
}

/**
 * Add a summary sheet to the workbook
 * @param workbook - The Excel workbook
 * @param data - The processed Katapult data
 */
function addSummarySheet(workbook: XLSX.WorkBook, data: any) {
  const connections = data.connections || [];
  const totalConnections = connections.length;
  const refConnections = connections.filter(c => c.isRefConnection).length;
  const regularConnections = totalConnections - refConnections;
  
  // Count wires by category
  const wireCounts = {
    [WireCategory.COMMUNICATION]: 0,
    [WireCategory.CPS_ELECTRICAL]: 0,
    [WireCategory.OTHER]: 0,
    total: 0
  };
  
  // Count proposed changes
  let proposedChanges = 0;
  
  connections.forEach(connection => {
    Object.values(connection.wires).forEach(wire => {
      if (wire.category) {
        wireCounts[wire.category]++;
        wireCounts.total++;
      }
      
      if (wire.finalProposedMidspanHeight !== null || wire.finalProposedPoleAttachmentHeight !== null) {
        proposedChanges++;
      }
    });
  });
  
  // Create summary data
  const summaryData = [
    ['Katapult Analysis Summary'],
    [],
    ['Total Connections', totalConnections],
    ['Regular Connections', regularConnections],
    ['REF Connections', refConnections],
    [],
    ['Total Wires', wireCounts.total],
    ['Communication Wires', wireCounts[WireCategory.COMMUNICATION]],
    ['CPS Electrical Wires', wireCounts[WireCategory.CPS_ELECTRICAL]],
    ['Other Wires', wireCounts[WireCategory.OTHER]],
    [],
    ['Wires with Proposed Changes', proposedChanges],
  ];
  
  // Add summary sheet
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');
}

/**
 * Add a connections sheet to the workbook
 * @param workbook - The Excel workbook
 * @param data - The processed Katapult data
 */
function addConnectionsSheet(workbook: XLSX.WorkBook, data: any) {
  const connections = data.connections || [];
  
  // Create connections header
  const connectionsHeader = [
    'Connection ID',
    'From Pole ID',
    'To Pole ID',
    'Connection Type',
    'Is REF',
    'Total Wires'
  ];
  
  // Create connections data
  const connectionsData = connections.map(connection => [
    connection.connectionId,
    connection.fromPoleId,
    connection.toPoleId || 'N/A',
    connection.buttonType,
    connection.isRefConnection ? 'Yes' : 'No',
    Object.keys(connection.wires).length
  ]);
  
  // Add connections sheet
  const connectionsSheet = XLSX.utils.aoa_to_sheet([
    connectionsHeader,
    ...connectionsData
  ]);
  
  XLSX.utils.book_append_sheet(workbook, connectionsSheet, 'Connections');
}

/**
 * Add a wire data sheet to the workbook
 * @param workbook - The Excel workbook
 * @param data - The processed Katapult data
 */
function addWireDataSheet(workbook: XLSX.WorkBook, data: any) {
  const connections = data.connections || [];
  
  // Create wire data header
  const wireDataHeader = [
    'Connection ID',
    'From Pole ID',
    'To Pole ID',
    'Trace ID',
    'Company',
    'Cable Type',
    'Category',
    'Lowest Midspan Height',
    'Proposed Midspan Height',
    'Has Proposed Change'
  ];
  
  // Create wire data
  const wireData = connections.flatMap(connection => 
    Object.values(connection.wires).map((wire: any) => {
      const hasProposedChange = wire.finalProposedMidspanHeight !== null && 
                               wire.lowestExistingMidspanHeight !== null && 
                               wire.finalProposedMidspanHeight !== wire.lowestExistingMidspanHeight;
      
      return [
        connection.connectionId,
        connection.fromPoleId,
        connection.toPoleId || 'N/A',
        wire.traceId,
        wire.company,
        wire.cableType,
        wire.category,
        formatHeightToString(wire.lowestExistingMidspanHeight),
        formatHeightToString(wire.finalProposedMidspanHeight),
        hasProposedChange ? 'Yes' : 'No'
      ];
    })
  );
  
  // Add wire data sheet
  const wireDataSheet = XLSX.utils.aoa_to_sheet([
    wireDataHeader,
    ...wireData
  ]);
  
  XLSX.utils.book_append_sheet(workbook, wireDataSheet, 'Wire Data');
}

/**
 * Add a REF connections sheet to the workbook
 * @param workbook - The Excel workbook
 * @param data - The processed Katapult data
 */
function addRefConnectionsSheet(workbook: XLSX.WorkBook, data: any) {
  const refConnections = data.connections.filter(c => c.isRefConnection) || [];
  
  if (refConnections.length === 0) {
    // If no REF connections, add empty sheet
    const emptySheet = XLSX.utils.aoa_to_sheet([
      ['No REF connections found in the data']
    ]);
    XLSX.utils.book_append_sheet(workbook, emptySheet, 'REF Connections');
    return;
  }
  
  // Create REF connections header
  const refConnectionsHeader = [
    'Connection ID',
    'Pole ID',
    'Trace ID',
    'Company',
    'Cable Type',
    'Category',
    'Attachment Height',
    'Proposed Attachment Height',
    'Has Proposed Change'
  ];
  
  // Create REF connections data
  const refConnectionsData = refConnections.flatMap(connection => 
    Object.values(connection.wires).map((wire: any) => {
      const hasProposedChange = wire.finalProposedPoleAttachmentHeight !== null && 
                               wire.lowestExistingPoleAttachmentHeight !== null && 
                               wire.finalProposedPoleAttachmentHeight !== wire.lowestExistingPoleAttachmentHeight;
      
      return [
        connection.connectionId,
        connection.fromPoleId,
        wire.traceId,
        wire.company,
        wire.cableType,
        wire.category,
        formatHeightToString(wire.lowestExistingPoleAttachmentHeight),
        formatHeightToString(wire.finalProposedPoleAttachmentHeight),
        hasProposedChange ? 'Yes' : 'No'
      ];
    })
  );
  
  // Add REF connections sheet
  const refConnectionsSheet = XLSX.utils.aoa_to_sheet([
    refConnectionsHeader,
    ...refConnectionsData
  ]);
  
  XLSX.utils.book_append_sheet(workbook, refConnectionsSheet, 'REF Connections');
}
