
import { useState } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { PoleData } from '@/services/poleDataProcessor';

interface DataPreviewProps {
  spidaData: any;
  katapultData: any;
  processedPoles: PoleData[];
}

const DataPreview = ({ spidaData, katapultData, processedPoles }: DataPreviewProps) => {
  const [viewMode, setViewMode] = useState<'raw' | 'processed'>('processed');

  const formatJson = (json: any) => {
    try {
      return JSON.stringify(json, null, 2);
    } catch (error) {
      return 'Invalid JSON';
    }
  };

  const formatProcessedData = (poles: PoleData[]) => {
    if (!poles.length) return 'No processed data available';
    
    const summaryItems = poles.map(pole => (
      `Pole #${pole.poleNumber} (${pole.operationNumber}): ${pole.attachmentAction}, Owner: ${pole.poleOwner}, PLA: ${pole.pla}, Spans: ${pole.spans.length}`
    )).join('\n');
    
    return summaryItems;
  };

  return (
    <Card className="w-full">
      <CardHeader className="pb-3">
        <CardTitle className="text-lg flex justify-between items-center">
          <span>Data Preview</span>
          <div className="flex items-center space-x-2">
            <Tabs defaultValue="processed" className="w-[240px]">
              <TabsList>
                <TabsTrigger value="processed" onClick={() => setViewMode('processed')}>Processed</TabsTrigger>
                <TabsTrigger value="raw" onClick={() => setViewMode('raw')}>Raw JSON</TabsTrigger>
              </TabsList>
            </Tabs>
          </div>
        </CardTitle>
      </CardHeader>
      <CardContent>
        {viewMode === 'raw' ? (
          <Tabs defaultValue="spida">
            <TabsList className="mb-4">
              <TabsTrigger value="spida">SPIDA Data</TabsTrigger>
              <TabsTrigger value="katapult">Katapult Data</TabsTrigger>
            </TabsList>
            <TabsContent value="spida">
              <pre className="overflow-auto max-h-96 bg-gray-50 p-4 rounded-md text-xs font-mono">
                {spidaData ? formatJson(spidaData) : 'No SPIDA data available'}
              </pre>
            </TabsContent>
            <TabsContent value="katapult">
              <pre className="overflow-auto max-h-96 bg-gray-50 p-4 rounded-md text-xs font-mono">
                {katapultData ? formatJson(katapultData) : 'No Katapult data available'}
              </pre>
            </TabsContent>
          </Tabs>
        ) : (
          <div className="max-h-96 overflow-auto bg-gray-50 p-4 rounded-md">
            <pre className="text-xs font-mono whitespace-pre-wrap">
              {processedPoles.length ? formatProcessedData(processedPoles) : 'No processed data available yet. Process files to see results.'}
            </pre>
          </div>
        )}
      </CardContent>
    </Card>
  );
};

export default DataPreview;
