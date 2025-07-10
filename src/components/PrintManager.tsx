import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Printer, AlertTriangle, Loader2 } from "lucide-react";
import { MultiFileUpload } from "./MultiFileUpload";
import { showSuccess, showError, showWarning } from "@/utils/toast";

export function PrintManager() {
  const [files, setFiles] = useState<File[]>([]);
  const [isPrinting, setIsPrinting] = useState(false);

  const handlePrint = async () => {
    if (files.length === 0) {
      showError("Per favore, carica almeno un file .pdf da stampare.");
      return;
    }

    setIsPrinting(true);
    showWarning(`Verrà aperta una finestra di dialogo di stampa per ogni file (${files.length} in totale).`);

    for (const file of files) {
      try {
        const url = URL.createObjectURL(file);
        const iframe = document.createElement('iframe');
        iframe.style.display = 'none';
        document.body.appendChild(iframe);

        await new Promise<void>((resolve, reject) => {
          iframe.onload = () => {
            setTimeout(() => {
              try {
                iframe.contentWindow?.focus();
                iframe.contentWindow?.print();
                resolve();
              } catch (e) {
                reject(e);
              } finally {
                URL.revokeObjectURL(url);
                document.body.removeChild(iframe);
              }
            }, 500); // Attende un po' per il rendering del PDF
          };
          iframe.onerror = (e) => {
            reject(e);
            URL.revokeObjectURL(url);
            document.body.removeChild(iframe);
          };
          iframe.src = url;
        });
        
        // Pausa per evitare che le finestre di dialogo si sovrappongano in modo aggressivo
        await new Promise(resolve => setTimeout(resolve, 1000));

      } catch (error) {
        console.error("Errore durante la stampa del file:", file.name, error);
        showError(`Impossibile stampare il file ${file.name}.`);
        break; // Interrompe il ciclo in caso di errore
      }
    }
    
    setIsPrinting(false);
    if (!isPrinting) { // Evita notifiche multiple
        showSuccess("Processo di stampa completato.");
    }
  };

  return (
    <Card>
      <CardHeader>
        <CardTitle>Stampa Documenti</CardTitle>
        <CardDescription>Carica i documenti PDF e avvia la stampa tramite il dialogo del browser.</CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <Alert variant="default">
          <AlertTriangle className="h-4 w-4" />
          <AlertTitle>Come funziona la stampa</AlertTitle>
          <AlertDescription>
            Per motivi di sicurezza, l'applicazione userà la finestra di dialogo di stampa del tuo browser. Potrai selezionare la stampante e le opzioni (copie, fronte-retro, ecc.) direttamente lì.
          </AlertDescription>
        </Alert>

        <MultiFileUpload
          id="pdf-files"
          label="Carica Documenti (.pdf)"
          files={files}
          onFilesChange={setFiles}
          accept=".pdf,application/pdf"
          helpText="Puoi selezionare più file PDF."
        />
        
        <Button size="lg" onClick={handlePrint} className="w-full" disabled={files.length === 0 || isPrinting}>
          {isPrinting ? <Loader2 className="mr-2 h-5 w-5 animate-spin" /> : <Printer className="mr-2 h-5 w-5" />}
          {isPrinting ? 'Stampa in corso...' : 'Stampa Documenti'}
        </Button>
      </CardContent>
    </Card>
  );
}