
/// <reference types="vite/client" />

interface Window {
  _extractMidspanHeightsOverride?: (katapultJson: any, spidaJson: any) => {
    comHeight: string;
    cpsHeight: string;
  };
}
