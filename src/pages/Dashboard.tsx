import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Label } from "@/components/ui/label";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Download, Info } from "lucide-react";
import { format } from "date-fns";
import { showError, showSuccess, showLoading, dismissToast } from "@/utils/toast";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { ResizableHandle, ResizablePanel, ResizablePanelGroup } from "@/components/ui/resizable";
import { Step } from "@/components/Step";
import { FileUpload } from "@/components/FileUpload";

import * as XLSX from "xlsx";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import { saveAs } from "file-saver";
import JSZip from "jszip";

const getCellValueAsString = (cellValue: any): string => {
  return cellValue !== null && cellValue !== undefined ? String(cellValue).trim() : '';
};

interface Discente {
  Matricola: string;
  "Grado militare": string;
  "Cognome e Nome del Discente": string;
  Cognome: string;
  Nome: string;
  Categoria: string;
}

interface CourseInfo {
  title: string;
  location: string;
  period: string;
  currentDate: string;
}

function Dashboard() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [wordFile, setWordFile] = useState<File | null>(null);
  const [discenti, setDiscenti] = useState<Discente[]>([]);
  const [signer, setSigner] = useState<string>("Il Direttore del Corso");
  const [courseInfo, setCourseInfo] = useState<CourseInfo | null>(null);

  const handleExcelUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setExcelFile(file);
    const reader = new FileReader();
    reader.onload = (e) => processExcelFile(e.target?.result as ArrayBuffer);
    reader.readAsArrayBuffer(file);
  };

  const processExcelFile = (data: ArrayBuffer) => {
    try {
      setDiscenti([]);
      setCourseInfo(null);
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];

      const title = getCellValueAsString(worksheet['A5']?.v);
      const location = getCellValueAsString(worksheet['C6']?.v);
      const startDate = worksheet['D6']?.v ? (worksheet['D6'].v instanceof Date ? format(worksheet['D6'].v, "dd/MM/yyyy") : getCellValueAsString(worksheet['D6'].v)) : "";
      const endDate = worksheet['E6']?.v ? (worksheet['E6'].v instanceof Date ? format(worksheet['E6'].v, "dd/MM/yyyy") : getCellValueAsString(worksheet['E6'].v)) : "";

      if (!title || !location || !startDate || !endDate) {
        showError("Dati corso mancanti o non validi nelle celle A5, C6, D6, E6. Controlla il file.");
        return;
      }
      const period = `dal ${startDate} al ${endDate}`;
      setCourseInfo({ title, location, period, currentDate: format(new Date(), "dd/MM/yyyy") });
      showSuccess("Dati del corso estratti con successo.");

      const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", blankrows: true });
      let headerRowIndex = -1;
      let headers: string[] = [];
      for (let i = 0; i < Math.min(jsonData.length, 20); i++) {
        const row = jsonData[i].map(h => getCellValueAsString(h).toLowerCase());
        if (row.some(h => h.includes('matricola')) && row.some(h => h.includes('grado')) && row.some(h => h.includes('cognome') || h.includes('nominativo'))) {
          headerRowIndex = i;
          headers = row;
          break;
        }
      }

      if (headerRowIndex === -1) {
        showError("Riga delle intestazioni non trovata. Assicurati che il file contenga 'Matricola', 'Grado' e 'Cognome'.");
        return;
      }

      const dataRows = jsonData.slice(headerRowIndex + 1);
      const findIndex = (keywords: string[]) => headers.findIndex(h => keywords.some(kw => h.includes(kw)));
      const matricolaIndex = findIndex(['matricola']);
      const gradoIndex = findIndex(['grado']);
      const categoriaIndex = findIndex(['cat.', 'cat', 'categoria']);
      const cognomeIndex = findIndex(['cognome']);
      const nomeIndex = findIndex(['nome']);
      const cognomeNomeIndex = findIndex(['cognome e nome', 'nominativo']);

      const mappedData = dataRows.map((row, rowIndex) => {
        if (row.every(cell => !getCellValueAsString(cell))) return null;
        let cognomeNome = '', cognome = '', nome = '';
        if (cognomeIndex !== -1 && nomeIndex !== -1 && getCellValueAsString(row[cognomeIndex])) {
          cognome = getCellValueAsString(row[cognomeIndex]);
          nome = getCellValueAsString(row[nomeIndex]);
          cognomeNome = `${cognome} ${nome}`.trim();
        } else if (cognomeNomeIndex !== -1 && getCellValueAsString(row[cognomeNomeIndex])) {
          cognomeNome = getCellValueAsString(row[cognomeNomeIndex]);
          const parts = cognomeNome.split(' ').filter(p => p);
          if (parts.length > 1) {
            nome = parts.pop() as string;
            cognome = parts.join(' ');
          } else {
            cognome = cognomeNome;
          }
        } else if (cognomeIndex !== -1) {
          cognome = getCellValueAsString(row[cognomeIndex]);
          cognomeNome = cognome;
        }
        const matricola = getCellValueAsString(row[matricolaIndex]);
        const grado = getCellValueAsString(row[gradoIndex]);
        if (!matricola || !grado || !cognomeNome) return null;
        return {
          Matricola: matricola,
          "Grado militare": grado,
          "Cognome e Nome del Discente": cognomeNome,
          Cognome: cognome,
          Nome: nome,
          Categoria: getCellValueAsString(row[categoriaIndex]),
          originalRow: headerRowIndex + 2 + rowIndex,
        };
      }).filter(Boolean) as (Discente & { originalRow: number })[];

      if (mappedData.some(d => !d.Nome)) {
        showError("Errore: Il campo Nome è vuoto o non è stato possibile estrarlo per alcuni discenti. Caricamento interrotto.");
        return;
      }
      setDiscenti(mappedData);
      showSuccess(`Caricamento completato. Trovati ${mappedData.length} discenti.`);
    } catch (error) {
      console.error("Errore durante l'elaborazione del file Excel:", error);
      showError("Errore imprevisto durante la lettura del file Excel.");
    }
  };

  const handleWordUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      setWordFile(file);
      showSuccess(`Template Word "${file.name}" caricato.`);
    }
  };

  const handleGenerateDocument = async () => {
    if (!wordFile || discenti.length === 0 || !courseInfo) {
      showError("Carica il file Excel e il template Word prima di generare i documenti.");
      return;
    }
    const toastId = showLoading("Generazione dei documenti in corso...");
    try {
      const content = await wordFile.arrayBuffer();
      const outputZip = new JSZip();
      for (const discente of discenti) {
        const templateZip = new PizZip(content);
        const doc = new Docxtemplater(templateZip, { paragraphLoop: true, linebreaks: true });
        doc.setData({
          titolocorso: courseInfo.title,
          categoria: discente.Categoria,
          localita: courseInfo.location,
          periodo_corso: courseInfo.period,
          firmatario: `${signer}\nCol. Massimiliano Fortino`,
          grado: discente["Grado militare"],
          cognome_nome: discente["Cognome e Nome del Discente"],
          cognome: discente.Cognome,
          nome: discente.Nome,
          matricola: discente.Matricola,
          datafirma: courseInfo.currentDate,
        });
        doc.render();
        const out = doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
        outputZip.file(`Attestato_${discente["Cognome e Nome del Discente"].replace(/[^a-zA-Z0-9]/g, '_')}.docx`, out);
      }
      const zipBlob = await outputZip.generateAsync({ type: "blob" });
      saveAs(zipBlob, "documenti_individuali_docx.zip");
      dismissToast(toastId);
      showSuccess("Archivio ZIP con documenti DOCX generato con successo!");
    } catch (error: any) {
      dismissToast(toastId);
      console.error("Errore nella generazione del documento:", error);
      showError(`Errore durante la generazione: ${error.message}`);
    }
  };

  return (
    <div className="container mx-auto p-4 h-screen flex flex-col">
      <header className="text-center mb-6">
        <h1 className="text-3xl font-bold tracking-tight">Dashboard Stampa Unione</h1>
        <p className="text-muted-foreground">Genera documenti personalizzati in pochi passaggi.</p>
      </header>

      <ResizablePanelGroup direction="vertical" className="flex-grow rounded-lg border">
        <ResizablePanel defaultSize={50} minSize={30}>
          <div className="p-6 space-y-8 h-full overflow-y-auto">
            <div className="grid md:grid-cols-2 gap-8">
              <div>
                <Step step={1} title="Carica i File" />
                <div className="space-y-4 mt-4">
                  <FileUpload
                    id="excel-file"
                    label="File Dati Discenti"
                    file={excelFile}
                    onUpload={handleExcelUpload}
                    onRemove={() => { setExcelFile(null); setDiscenti([]); setCourseInfo(null); }}
                    accept=".xlsx"
                    helpText="Carica il file Excel con i dati dei discenti."
                  />
                  <FileUpload
                    id="word-file"
                    label="Template Documento"
                    file={wordFile}
                    onUpload={handleWordUpload}
                    onRemove={() => setWordFile(null)}
                    accept=".docx"
                    helpText="Carica il template Word (.docx) con i segnaposto."
                  />
                </div>
              </div>
              <div>
                <Step step={2} title="Scegli il Firmatario" />
                <div className="mt-4">
                  <RadioGroup value={signer} onValueChange={setSigner} className="space-y-2">
                    <div className="flex items-center space-x-2"><RadioGroupItem value="Il Direttore del Corso" id="r1" /><Label htmlFor="r1">Il Direttore del Corso</Label></div>
                    <div className="flex items-center space-x-2"><RadioGroupItem value="Il Comandante del Centro" id="r2" /><Label htmlFor="r2">Il Comandante del Centro</Label></div>
                  </RadioGroup>
                  <p className="text-sm text-muted-foreground mt-2">Il titolo apparirà sopra: <strong>Col. Massimiliano Fortino</strong>.</p>
                </div>
              </div>
            </div>
          </div>
        </ResizablePanel>
        <ResizableHandle withHandle />
        <ResizablePanel defaultSize={50} minSize={30}>
          <div className="p-6 h-full flex flex-col">
            <Step step={3} title="Anteprima e Generazione" />
            {courseInfo && (
              <Card className="my-6">
                <CardHeader><CardTitle className="flex items-center gap-2 text-base"><Info size={16} /> Dati del Corso</CardTitle></CardHeader>
                <CardContent className="text-sm space-y-1">
                  <p><strong>Titolo:</strong> {courseInfo.title}</p>
                  <p><strong>Sede:</strong> {courseInfo.location}</p>
                  <p><strong>Periodo:</strong> {courseInfo.period}</p>
                </CardContent>
              </Card>
            )}
            {discenti.length > 0 ? (
              <div className="flex-grow flex flex-col min-h-0">
                <p className="mb-2"><strong>Discenti Trovati:</strong> <span className="font-mono p-1 bg-muted rounded-md">{discenti.length}</span></p>
                <div className="flex-grow overflow-y-auto rounded-md border">
                  <Table>
                    <TableHeader><TableRow><TableHead>Matricola</TableHead><TableHead>Grado</TableHead><TableHead>Cognome</TableHead><TableHead>Nome</TableHead></TableRow></TableHeader>
                    <TableBody>
                      {discenti.map((d, index) => (
                        <TableRow key={index}><TableCell>{d.Matricola}</TableCell><TableCell>{d["Grado militare"]}</TableCell><TableCell>{d.Cognome}</TableCell><TableCell>{d.Nome}</TableCell></TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>
              </div>
            ) : (
              <div className="flex-grow flex items-center justify-center text-center text-muted-foreground bg-muted/50 rounded-md">
                <p>I dati dei discenti appariranno qui dopo il caricamento del file Excel.</p>
              </div>
            )}
            <div className="mt-auto pt-6 flex gap-4">
              <Button size="lg" onClick={handleGenerateDocument} className="w-full" disabled={!wordFile || discenti.length === 0}>
                <Download className="mr-2 h-5 w-5" /> Genera DOCX e ZIP
              </Button>
            </div>
          </div>
        </ResizablePanel>
      </ResizablePanelGroup>
    </div>
  );
}

export default Dashboard;