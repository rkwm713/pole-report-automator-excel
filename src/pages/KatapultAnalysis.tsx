
import { useState } from 'react';
import { Button } from '@/components/ui/button';
import { toast } from 'sonner';
import { Separator } from '@/components/ui/separator';
import FileUploader from '@/components/FileUploader';
import ProcessingStatus from '@/components/ProcessingStatus';
import KatapultDataVisualization from '@/components/KatapultDataVisualization';
import { processKatapultData } from '@/utils/katapultDataProcessor';
import { downloadKatapultAnalysisExcel } from '@/services/katapultDataAnalyzer';

const KatapultAnalysis = () => {
  const [katapultFile, setKatapultFile] = useState<File | null>(null);
  const [katapultData, setKatapultData] = useState<any>(null);
  const [processedData, setProcessedData] = useState<any>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isComplete, setIsComplete] = useState(false);
  const [hasError, setHasError] = useState(false);
  const [errorDetails, setErrorDetails] = useState<string | undefined>(undefined);
  const [progress, setProgress] = useState(0);

  const handleFileSelected = (file: File, fileType: 'spida' | 'katapult') => {
    if (fileType === 'katapult') {
      setKatapultFile(file);
    }
  };

  const handleProcessData = async () => {
    if (!katapultFile) {
      toast.error('Missing required file', {
        description: 'Please upload the Katapult JSON file.'
      });
      return;
    }

    setIsProcessing(true);
    setIsComplete(false);
    setHasError(false);
    setErrorDetails(undefined);
    setProgress(0);

    try {
      // Read the Katapult file
      const katapultJson = await readJsonFile(katapultFile);
      setKatapultData(katapultJson);
      setProgress(50);

      // Process Katapult data
      const results = processKatapultData(katapultJson);
      setProcessedData(results);
      setProgress(100);
      
      setIsComplete(true);
      toast.success('Processing complete!');
    } catch (error) {
      console.error('Error processing data:', error);
      setHasError(true);
      setErrorDetails(error instanceof Error ? error.message : 'Unknown error occurred');
      toast.error('Processing failed', {
        description: 'An error occurred while processing the files.'
      });
    } finally {
      setIsProcessing(false);
    }
  };

  const handleReset = () => {
    setKatapultFile(null);
    setKatapultData(null);
    setProcessedData(null);
    setIsProcessing(false);
    setIsComplete(false);
    setHasError(false);
    setErrorDetails(undefined);
    setProgress(0);
    toast.info('Analysis reset');
  };

  const handleGenerateExcel = () => {
    if (processedData) {
      downloadKatapultAnalysisExcel(processedData, katapultData);
      toast.success('Excel report generated');
    }
  };

  const readJsonFile = (file: File): Promise<any> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const json = JSON.parse(e.target?.result as string);
          resolve(json);
        } catch (error) {
          reject(new Error('Invalid JSON file format'));
        }
      };
      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsText(file);
    });
  };

  return (
    <div className="container py-8">
      <header className="mb-8">
        <h1 className="text-3xl font-bold mb-2">Katapult Analysis</h1>
        <p className="text-muted-foreground">
          Process and analyze Katapult JSON data to extract midspan heights and pole attachment information.
        </p>
      </header>
      
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8 mb-8">
        <div className="lg:col-span-1">
          <FileUploader
            label="Katapult JSON File"
            fileType="katapult"
            onFileSelected={handleFileSelected}
            isUploaded={!!katapultFile}
          />

          <div className="mt-6 flex flex-col gap-3">
            <Button
              onClick={handleProcessData}
              disabled={isProcessing || !katapultFile}
              className="w-full"
            >
              Process Data
            </Button>
            
            <Button
              variant="outline"
              onClick={handleReset}
              disabled={isProcessing}
              className="w-full"
            >
              Reset
            </Button>
          </div>
        </div>
        
        <div className="lg:col-span-2">
          <ProcessingStatus
            isProcessing={isProcessing}
            isComplete={isComplete}
            hasError={hasError}
            progress={progress}
            onGenerateExcel={handleGenerateExcel}
            onReset={handleReset}
            errorDetails={errorDetails}
            poleCount={processedData?.connections?.length || 0}
          />
        </div>
      </div>

      <Separator className="my-8" />
      
      {isComplete && processedData && (
        <KatapultDataVisualization data={processedData} />
      )}
    </div>
  );
};

export default KatapultAnalysis;
