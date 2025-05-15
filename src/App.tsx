
import { Toaster } from "@/components/ui/toaster";
import { Toaster as Sonner } from "@/components/ui/sonner";
import { TooltipProvider } from "@/components/ui/tooltip";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { BrowserRouter, Routes, Route, Link } from "react-router-dom";
import Index from "./pages/Index";
import KatapultAnalysis from "./pages/KatapultAnalysis";
import NotFound from "./pages/NotFound";
import "./utils/poleDataOverrides"; // Import our overrides

const queryClient = new QueryClient();

const App = () => (
  <QueryClientProvider client={queryClient}>
    <TooltipProvider>
      <Toaster />
      <Sonner />
      <BrowserRouter>
        <nav className="bg-gray-100 py-4">
          <div className="container mx-auto flex gap-6">
            <Link to="/" className="text-blue-600 hover:text-blue-800">Home</Link>
            <Link to="/katapult" className="text-blue-600 hover:text-blue-800">Katapult Analysis</Link>
          </div>
        </nav>
        <Routes>
          <Route path="/" element={<Index />} />
          <Route path="/katapult" element={<KatapultAnalysis />} />
          {/* ADD ALL CUSTOM ROUTES ABOVE THE CATCH-ALL "*" ROUTE */}
          <Route path="*" element={<NotFound />} />
        </Routes>
      </BrowserRouter>
    </TooltipProvider>
  </QueryClientProvider>
);

export default App;
