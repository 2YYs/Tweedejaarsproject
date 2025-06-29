
import React from 'react';
import { ScrollArea } from "@/components/ui/scroll-area";
import { Badge } from "@/components/ui/badge";
import { CheckCircle2, AlertCircle, Info, Clock } from 'lucide-react';

interface ProcessingLog {
  timestamp: string;
  action: string;
  details: string;
  status: 'success' | 'error' | 'info';
}

interface LogViewerProps {
  logs: ProcessingLog[];
}

const LogViewer: React.FC<LogViewerProps> = ({ logs }) => {
  const getStatusIcon = (status: ProcessingLog['status']) => {
    switch (status) {
      case 'success':
        return <CheckCircle2 className="w-4 h-4 text-green-600" />;
      case 'error':
        return <AlertCircle className="w-4 h-4 text-red-600" />;
      case 'info':
        return <Info className="w-4 h-4 text-blue-600" />;
      default:
        return <Clock className="w-4 h-4 text-gray-600" />;
    }
  };

  const getStatusBadge = (status: ProcessingLog['status']) => {
    switch (status) {
      case 'success':
        return <Badge className="bg-green-100 text-green-800 border-green-200">Succes</Badge>;
      case 'error':
        return <Badge className="bg-red-100 text-red-800 border-red-200">Fout</Badge>;
      case 'info':
        return <Badge className="bg-blue-100 text-blue-800 border-blue-200">Info</Badge>;
      default:
        return <Badge variant="outline">Onbekend</Badge>;
    }
  };

  if (logs.length === 0) {
    return (
      <div className="text-center py-8 text-gray-500">
        <Clock className="w-8 h-8 mx-auto mb-2 text-gray-400" />
        <p className="text-sm">Nog geen activiteiten</p>
        <p className="text-xs text-gray-400 mt-1">
          Logs verschijnen hier tijdens het verwerken
        </p>
      </div>
    );
  }

  return (
    <ScrollArea className="h-[400px] pr-4">
      <div className="space-y-3">
        {logs.map((log, index) => (
          <div
            key={index}
            className="border rounded-lg p-3 bg-white shadow-sm hover:shadow-md transition-shadow duration-200"
          >
            <div className="flex items-start space-x-3">
              <div className="flex-shrink-0 mt-1">
                {getStatusIcon(log.status)}
              </div>
              
              <div className="flex-1 min-w-0">
                <div className="flex items-center justify-between mb-1">
                  <p className="text-sm font-medium text-gray-900 truncate">
                    {log.action}
                  </p>
                  {getStatusBadge(log.status)}
                </div>
                
                <p className="text-xs text-gray-600 mb-2">
                  {log.details}
                </p>
                
                <div className="flex items-center space-x-1 text-xs text-gray-500">
                  <Clock className="w-3 h-3" />
                  <span>{log.timestamp}</span>
                </div>
              </div>
            </div>
          </div>
        ))}
      </div>
    </ScrollArea>
  );
};

export default LogViewer;
