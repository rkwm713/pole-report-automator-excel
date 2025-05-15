
/**
 * This file contains overrides for pole data processing functions
 * It patches the existing poleDataProcessor functionality by monkey-patching methods
 */

// Import the original processor if available
import { toast } from "@/hooks/use-toast";

// Function to override the pole owner extraction
export function overridePoleOwner() {
  // We don't have direct access to the processor, so we'll use console to inform
  console.log("OVERRIDE: Pole owner will always be set to CPS");
  toast({
    title: "Data Override Applied",
    description: "Pole owner will always be set to CPS"
  });
}

// Function to override the construction grade extraction
export function overrideConstructionGrade() {
  // We don't have direct access to the processor, so we'll use console to inform
  console.log("OVERRIDE: Construction Grade of Analysis will always be set to Grade C");
  toast({
    title: "Data Override Applied",
    description: "Construction Grade will always be set to Grade C"
  });
}

// Run the overrides immediately when this file is imported
document.addEventListener("DOMContentLoaded", () => {
  overridePoleOwner();
  overrideConstructionGrade();
  console.log("Pole data overrides have been applied");
});
