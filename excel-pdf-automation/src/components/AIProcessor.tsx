import { useState } from 'react';
import { Button } from '@/components/ui/button';
import { Loader2, RotateCcw, Download, Brain } from 'lucide-react';
import { toast } from 'sonner';

// Gebruik de omgevingsvariabele voor de API URL, met localhost als fallback voor lokale ontwikkeling.
const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:5001';

interface AIProcessorProps {
  pdfFiles: File[];
  onProcessingComplete: () => void;
  onReset: () => void;
  addLog: (action: string, details: string, status?: 'success' | 'error' | 'info') => void;
}

const AIProcessor = ({ pdfFiles, onProcessingComplete, onReset, addLog }: AIProcessorProps) => {
  const [isProcessing, setIsProcessing] = useState(false);
  const [isComplete, setIsComplete] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);

  const triggerDownload = (url?: string) => {
    const finalUrl = url || downloadUrl;
    if (!finalUrl) {
      toast.error("Downloadfout", { description: "De downloadlink is niet beschikbaar." });
      return;
    }
    const link = document.createElement('a');
    link.href = finalUrl;
    link.setAttribute('download', 'cao_samenvatting_resultaat.xlsx');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    // URL.revokeObjectURL(finalUrl); // Clean up the object URL after download
  };

  const startProcessing = async () => {
    setIsProcessing(true);
    addLog('Verwerking Gestart', `Bezig met het versturen van ${pdfFiles.length} bestand(en) naar de server.`, 'info');

    const formData = new FormData();
    pdfFiles.forEach(file => {
      formData.append('files', file);
    });

    try {
      addLog('Communicatie', 'Wachten op antwoord van de server...', 'info');
      const response = await fetch(`${API_BASE_URL}/api/process`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        // Probeer de serverfout te lezen
        const errorData = await response.json().catch(() => ({ error: 'Onbekende serverfout.' }));
        const errorMessage = errorData.error || `HTTP status ${response.status}`;
        throw new Error(errorMessage);
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url); // Bewaar de URL voor 'Download Opnieuw'

      addLog('Succes', 'Excel-bestand succesvol ontvangen van de server.', 'success');
      toast.success("Verwerking voltooid!", {
        description: "Het Excel-bestand wordt nu gedownload.",
      });

      // Start de download automatisch
      triggerDownload(url);

      setIsComplete(true);
      onProcessingComplete();

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      toast.error("Verwerkingsfout", {
        description: `Er is een fout opgetreden: ${errorMessage}`,
      });
      addLog('Fout', `Er is een fout opgetreden: ${errorMessage}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  if (isComplete) {
    return (
      <div className="flex flex-col items-center justify-center p-8 border-2 border-dashed rounded-lg bg-green-50">
        <Download className="w-10 h-10 text-green-600 mb-4" />
        <h3 className="text-lg font-semibold text-green-800">Verwerking Voltooid</h3>
        <p className="text-sm text-green-700 mb-6 text-center">
          De download is gestart. Klik hieronder als het downloaden niet automatisch begon.
        </p>
        <div className="flex items-center space-x-4">
          <Button onClick={() => triggerDownload()}>
            <Download className="w-4 h-4 mr-2" />
            Download Opnieuw
          </Button>
          <Button onClick={onReset} variant="outline">
            <RotateCcw className="w-4 h-4 mr-2" />
            Opnieuw Beginnen
          </Button>
        </div>
      </div>
    );
  }

  return (
    <div className="text-center">
      <p className="text-sm text-gray-600 mb-4">
        {pdfFiles.length} PDF-bestand(en) zijn klaar voor analyse. Klik op 'Start Verwerking' om de AI aan het werk te zetten.
      </p>
      <Button 
        onClick={startProcessing} 
        disabled={isProcessing}
        className="bg-gradient-to-r from-blue-600 to-purple-600 hover:from-blue-700 hover:to-purple-700 text-white px-8 py-3 rounded-lg font-medium shadow-lg hover:shadow-xl transition-all duration-300 w-full"
      >
        {isProcessing ? (
          <>
            <Loader2 className="mr-2 h-5 w-5 animate-spin" />
            Verwerken...
          </>
        ) : (
          <>
            <Brain className="w-5 h-5 mr-2" />
            Start Analyse & Genereer Excel
          </>
        )}
      </Button>
    </div>
  );
};

export default AIProcessor;
