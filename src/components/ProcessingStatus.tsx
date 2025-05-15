
import { Progress } from '@/components/ui/progress';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { CircleCheckBig, CircleXIcon, FileSpreadsheet, Loader } from 'lucide-react';

export interface ProcessingStatusProps {
  isProcessing: boolean;
  isComplete: boolean;
  hasError: boolean;
  progress: number;
  onGenerateExcel: () => void;
  onReset: () => void;
  errorDetails?: string;
  poleCount?: number;
}

const ProcessingStatus = ({ 
  isProcessing, 
  isComplete, 
  hasError, 
  progress, 
  onGenerateExcel, 
  onReset,
  errorDetails,
  poleCount = 0
}: ProcessingStatusProps) => {

  const getStatusDisplay = () => {
    if (hasError) {
      return (
        <div className="flex flex-col items-center py-6">
          <CircleXIcon className="h-16 w-16 text-destructive mb-4" />
          <h3 className="text-lg font-medium mb-2">Processing Failed</h3>
          <p className="text-sm text-muted-foreground text-center mb-4">
            An error occurred during processing
          </p>
          {errorDetails && (
            <div className="w-full bg-destructive/10 p-3 rounded-md border border-destructive/20 mb-4">
              <p className="text-sm text-destructive font-mono whitespace-pre-wrap">
                {errorDetails}
              </p>
            </div>
          )}
          <Button variant="outline" onClick={onReset}>
            Try Again
          </Button>
        </div>
      );
    }
    
    if (isComplete) {
      return (
        <div className="flex flex-col items-center py-6">
          <CircleCheckBig className="h-16 w-16 text-green-500 mb-4" />
          <h3 className="text-lg font-medium mb-2">Processing Complete</h3>
          <p className="text-sm text-muted-foreground text-center mb-4">
            Successfully processed data for {poleCount} poles
          </p>
          <div className="flex gap-3">
            <Button 
              className="gap-2" 
              onClick={onGenerateExcel}
            >
              <FileSpreadsheet className="h-4 w-4" />
              Download Excel Report
            </Button>
            <Button variant="outline" onClick={onReset}>
              Reset
            </Button>
          </div>
        </div>
      );
    }
    
    return (
      <div className="flex flex-col items-center py-6">
        <div className="relative h-16 w-16 mb-4">
          <Loader className="h-16 w-16 text-blue-500 animate-spin" />
        </div>
        <h3 className="text-lg font-medium mb-2">Processing Data</h3>
        <p className="text-sm text-muted-foreground text-center mb-4">
          Analyzing and compiling report data...
        </p>
        <div className="w-full mb-4">
          <Progress value={progress} className="h-2" />
          <p className="text-xs text-center mt-1 text-muted-foreground">
            {progress}% Complete
          </p>
        </div>
      </div>
    );
  };
  
  return (
    <Card className="w-full">
      <CardHeader className="pb-3">
        <CardTitle className="text-lg">Processing Status</CardTitle>
      </CardHeader>
      <CardContent>
        {getStatusDisplay()}
      </CardContent>
    </Card>
  );
};

export default ProcessingStatus;
