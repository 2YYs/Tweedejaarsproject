import React, { useState } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { CheckCircle2, FileText, Cpu, Download, FileSpreadsheet, Plus } from 'lucide-react';
import PDFUpload from '@/components/PDFUpload';
import AIProcessor from '@/components/AIProcessor';
import LogViewer from '@/components/LogViewer';
import './mobile.css';
import dotsIcon from '@/assets/dots.svg';

interface ProcessingLog {
  timestamp: string;
  action: string;
  details: string;
  status: 'success' | 'error' | 'info';
}

const Index = () => {
  const [currentStep, setCurrentStep] = useState(1);
  const [pdfFiles, setPDFFiles] = useState<File[]>([]);
  const [isProcessingComplete, setIsProcessingComplete] = useState(false);
  const [logs, setLogs] = useState<ProcessingLog[]>([]);

  const addLog = (action: string, details: string, status: 'success' | 'error' | 'info' = 'info') => {
    const log: ProcessingLog = {
      timestamp: new Date().toLocaleString(),
      action,
      details,
      status
    };
    setLogs(prev => [...prev, log]);
  };

  const steps = [
    { number: 1, title: 'Upload CAO PDF', icon: FileText },
    { number: 2, title: 'AI Extractie', icon: Cpu },
    { number: 3, title: 'Nieuwe Rij Toevoegen', icon: Plus },
    { number: 4, title: 'Download Resultaat', icon: Download }
  ];

  const handlePDFUpload = (files: File[]) => {
    setPDFFiles(files);
    addLog('CAO PDFs Upload', `${files.length} PDF bestand(en) succesvol geüpload.`, 'success');
    setCurrentStep(2);
  };

  const handleProcessingComplete = () => {
    setIsProcessingComplete(true);
    addLog('Verwerking Voltooid', 'Het Excel-bestand is gegenereerd en wordt gedownload.', 'success');
  };

  const handleReset = () => {
    setCurrentStep(1);
    setPDFFiles([]);
    setIsProcessingComplete(false);
    setLogs([]);
    addLog('Reset', 'De applicatie is gereset.', 'info');
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-indigo-50 overflow-hidden relative">
      <img src={dotsIcon} alt="dots" style={{zIndex: '0', top: '-250px', opacity: '.6', width: '60%', height: '100%', left: '50%', position: 'absolute', transform: 'translateX(-50%)'}} />
      <div className="container mx-auto px-4 py-36 relative z-10">
        <div className="text-center mb-24">
          <h1 className="text-3xl font-semibold text-gray-900 mb-4">
            CAO Analyse & Excel Generator
          </h1>
          <p className="text-base text-gray-600 max-w-3xl mx-auto">
            Upload uw master Excel-bestand en CAO PDF om automatisch nieuwe rijen toe te voegen met geëxtraheerde gegevens.
          </p>
        </div>

        {/* Progress Steps */}
        <div className="flex justify-center mb-20">
          <div style={{paddingTop: '.8rem', paddingBottom: '.8rem'}} className="flex items-center space-x-4 bg-white rounded-lg shadow-sm p-6">
            {steps.map((step, index) => {
              const Icon = step.icon;
              
              const isCompleted = 
                (step.number === 1 && pdfFiles.length > 0) ||
                (step.number > 1 && step.number < 4 && isProcessingComplete);

              const isActive = 
                (step.number === 1 && pdfFiles.length === 0) ||
                (step.number === 2 && pdfFiles.length > 0 && !isProcessingComplete) ||
                (step.number === 4 && isProcessingComplete);

              return (
                <div key={step.number} className="flex items-center">
                  <div className="flex items-center space-x-2">
                    <div className={`flex items-center justify-center w-10 h-10 rounded-full border-2 transition-all duration-300 ${
                      isCompleted 
                        ? 'bg-green-500 border-green-500 text-white' 
                        : isActive 
                          ? 'bg-blue-500 border-blue-500 text-white' 
                          : 'border-gray-300 text-gray-400'
                    }`}>
                      {isCompleted ? (
                        <CheckCircle2 size={20} />
                      ) : (
                        <Icon size={20} />
                      )}
                    </div>
                    <span 
                      className={`mobile text-sm font-medium ${
                      isActive ? 'text-blue-600' : isCompleted ? 'text-green-600' : 'text-gray-400'
                    }`}>
                      {step.title}
                    </span>
                  </div>
                  {index < steps.length - 1 && (
                    <div className={`w-8 h-0.5 ${
                      isCompleted ? 'bg-green-500' : 'bg-gray-300'
                    }`} />
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* Main Content */}
        <div className="grid lg:grid-cols-3 gap-8">
          <div className="lg:col-span-2 space-y-6">
            {/* Step 1: PDF Upload */}
            <Card className={`transition-all duration-300 ${currentStep === 1 ? 'ring-2 ring-blue-500' : ''}`}>
              <CardHeader>
                <CardTitle className="flex items-center space-x-2">
                  <FileText className="w-5 h-5 text-blue-600" />
                  <span className='text-base'>Stap 1: Upload CAO PDF Bestanden</span>
                  {pdfFiles.length > 0 && <Badge variant="secondary">Voltooid</Badge>}
                </CardTitle>
              </CardHeader>
              <CardContent className="flex flex-col">
                <PDFUpload onUpload={handlePDFUpload} />
              </CardContent>
            </Card>

            {/* Step 2: AI Processing & Download */}
            {currentStep === 2 && pdfFiles.length > 0 && (
              <Card className={`transition-all duration-300 ring-2 ring-blue-500`}>
                <CardHeader>
                  <CardTitle className="flex items-center space-x-2">
                    <Cpu className="w-5 h-5 text-blue-600" />
                    <span className='text-base'>Stap 2: Verwerk & Download</span>
                    {isProcessingComplete && <Badge variant="secondary">Voltooid</Badge>}
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <AIProcessor 
                    pdfFiles={pdfFiles}
                    onProcessingComplete={handleProcessingComplete}
                    onReset={handleReset}
                    addLog={addLog}
                  />
                </CardContent>
              </Card>
            )}

          </div>

          {/* Log Viewer */}
          <div className="lg:col-span-1">
            <LogViewer logs={logs} />
          </div>
        </div>
      </div>
    </div>
  );
};

export default Index;
