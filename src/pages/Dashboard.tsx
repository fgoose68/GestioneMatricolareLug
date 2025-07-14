import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Label } from "@/components/ui/label";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Download, Info } from "lucide-react";
import { format } from "date-fns";
import { showError, showSuccess, showLoading, dismissToast } from "@/utils/toast";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { FileUpload } from "@/components/FileUpload";
import { PrintManager } from "@/components/PrintManager";
import { Switch } from "@/components/ui/switch";

import * as XLSX from "xlsx";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import { saveAs } from "file-saver";
import JSZip from "jszip";
import html2pdf from "html2pdf.js";
import ReactDOMServer from 'react-dom/server';
import { AttestatoTemplate } from '@/components/AttestatoTemplate';

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
  const [outputFormat, setOutputFormat] = useState<'docx' | 'pdf'>('pdf');

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

      if (!title || !location || !startDate) { 
        showError("Dati corso mancanti o non validi nelle celle A5, C6, D6. Controlla il file.");
        return;
      }
      
      let period = "";
      if (endDate) {
        period = `dal ${startDate} al ${endDate}`;
      } else {
        period = `il ${startDate}`; 
      }

      setCourseInfo({ title, location, period, currentDate: format(new Date(), "dd/MM/yyyy") });
      showSuccess("Dati del corso estratti con successo.");

      const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", blankrows: true });
      let headerRowIndex = -1;
      let headers: string[] = [];
      for (let i = 0; i < Math.min(jsonData.length, 20); i++) {
        const row = jsonData[i].map(h => getCellValueAsString(h).toLowerCase());
        const hasMatricola = row.some(h => h.includes('matricola'));
        const hasGrado = row.some(h => h.includes('grado'));
        const hasFullName = row.some(h => h.includes('cognome e nome') || h.includes('nominativo'));
        const hasSeparateNames = row.some(h => h.includes('cognome')) && row.some(h => h.includes('nome'));

        if (hasMatricola && hasGrado && (hasFullName || hasSeparateNames)) {
          headerRowIndex = i;
          headers = row;
          break;
        }
      }

      if (headerRowIndex === -1) {
        showError("Riga intestazioni non trovata. Assicurati che il file contenga 'Matricola', 'Grado', e ('Cognome' e 'Nome' o 'Nominativo').");
        return;
      }

      const dataRows = jsonData.slice(headerRowIndex + 1);
      const findIndex = (keywords: string[]) => headers.findIndex(h => keywords.some(kw => h.includes(kw)));
      
      const matricolaIndex = findIndex(['matricola']);
      const gradoIndex = findIndex(['grado']);
      const categoriaIndex = findIndex(['cat.', 'cat', 'categoria']);
      const cognomeIndex = findIndex(['cognome', 'cognome del discente', 'cognome discente', 'last name', 'surname']);
      const nomeIndex = findIndex(['nome', 'nome del discente', 'nome discente', 'first name']);
      const cognomeNomeIndex = findIndex(['cognome e nome', 'nominativo', 'cognome nome', 'full name', 'nominativo del discente']);

      const mappedData = dataRows.map((row, rowIndex) => {
        if (row.every(cell => !getCellValueAsString(cell))) return null;
        
        let cognome = '';
        let nome = '';

        if (cognomeIndex !== -1 && nomeIndex !== -1 && cognomeIndex !== nomeIndex) {
            cognome = getCellValueAsString(row[cognomeIndex]);
            nome = getCellValueAsString(row[nomeIndex]);
        } 
        else if (cognomeNomeIndex !== -1) {
            const fullName = getCellValueAsString(row[cognomeNomeIndex]);
            const parts = fullName.split(' ').filter(p => p);
            if (parts.length > 1) {
                nome = parts.pop() as string;
                cognome = parts.join(' ');
            } else {
                cognome = fullName;
            }
        }
        else if (cognomeIndex !== -1) {
            const fullName = getCellValueAsString(row[cognomeIndex]);
            const parts = fullName.split(' ').filter(p => p);
            if (parts.length > 1) {
                nome = parts.pop() as string;
                cognome = parts.join(' ');
            } else {
                cognome = fullName;
            }
        }

        const cognomeNome = `${cognome} ${nome}`.trim();
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

      if (mappedData.some(d => !d.Nome && !d.Cognome.includes(' '))) {
        showError("Avviso: Il campo Nome è vuoto per alcuni discenti. Controlla le colonne 'Cognome' e 'Nome' nel file Excel.");
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
    if ((outputFormat === 'docx' && !wordFile) || !excelFile || discenti.length === 0 || !courseInfo) {
      showError("Per generare i documenti, carica il file Excel. Per il formato DOCX, è necessario anche il template Word.");
      return;
    }
    const toastId = showLoading(`Generazione dei documenti in formato ${outputFormat.toUpperCase()} in corso...`);
    
    try {
      const outputZip = new JSZip();
      const docxContent = outputFormat === 'docx' && wordFile ? await wordFile.arrayBuffer() : null;

      for (const discente of discenti) {
        const fileName = `Attestato_${discente["Cognome e Nome del Discente"].replace(/[^a-zA-Z0-9]/g, '_')}`;

        let categoriaPerTemplate = discente.Categoria;
        const gradoMilitare = discente["Grado militare"].toUpperCase();

        if (['GCA', 'GDV', 'GDB', 'COL', 'TCL', 'MAG', 'CAP', 'TEN', 'STN'].includes(gradoMilitare)) {
          categoriaPerTemplate = "l'Ufficiale";
        } else if (['LGT.CS', 'LGT', 'MAR.A', 'MAR.C', 'MAR.O', 'MAR'].includes(gradoMilitare)) {
          categoriaPerTemplate = "l'Ispettore";
        } else if (['B.C.QS', 'BRIG.C', 'BRIG', 'VBRIG'].includes(gradoMilitare)) {
          categoriaPerTemplate = "il Sovrintendente";
        } else if (['APS.QS', 'APP.SC', 'APP', 'FIN.SC', 'FIN'].includes(gradoMilitare)) {
          categoriaPerTemplate = "il Militare";
        } else {
          categoriaPerTemplate = discente.Categoria;
        }

        if (outputFormat === 'pdf') {
          const componentHtml = ReactDOMServer.renderToString(
            <AttestatoTemplate
              courseInfo={courseInfo}
              discente={discente}
              signer={signer}
              categoria={categoriaPerTemplate}
            />
          );
          
          const pdfBlob = await html2pdf().from(componentHtml).set({
            margin: 0,
            filename: `${fileName}.pdf`,
            html2canvas: { scale: 2, useCORS: true },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
          }).output('blob');

          outputZip.file(`${fileName}.pdf`, pdfBlob);

        } else { // outputFormat is 'docx'
          if (!docxContent) continue;
          const templateZip = new PizZip(docxContent);
          const doc = new Docxtemplater(templateZip, { paragraphLoop: true, linebreaks: true });

          doc.setData({
            titolocorso: courseInfo.title,
            categoria: categoriaPerTemplate,
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
          outputZip.file(`${fileName}.docx`, out);
        }
      }

      const zipFileName = `documenti_${outputFormat}.zip`;
      const zipBlob = await outputZip.generateAsync({ type: "blob" });
      saveAs(zipBlob, zipFileName);
      dismissToast(toastId);
      showSuccess(`Archivio ZIP con documenti ${outputFormat.toUpperCase()} generato con successo!`);
    } catch (error: any) {
      dismissToast(toastId);
      console.error("Errore nella generazione del documento:", error);
      showError(`Errore durante la generazione: ${error.message}`);
    }
  };

  return (
    <div className="container mx-auto p-4 md:p-6 lg:p-8">
      <header className="text-center mb-8">
        <h1 className="text-3xl font-bold tracking-tight">Gestione Modelli L</h1>
        <p className="text-muted-foreground">Genera documenti personalizzati in pochi passaggi.</p>
      </header>

      <div className="space-y-8">
        <Card>
          <CardHeader>
            <CardTitle>Passo 1 & 2: Configurazione</CardTitle>
            <CardDescription>Carica i file necessari e scegli il firmatario per i documenti.</CardDescription>
          </CardHeader>
          <CardContent>
            <div className="grid md:grid-cols-2 gap-8">
              <div className="space-y-4">
                <h3 className="text-lg font-semibold">Carica i File</h3>
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
                  label="Template Documento (per .DOCX)"
                  file={wordFile}
                  onUpload={handleWordUpload}
                  onRemove={() => setWordFile(null)}
                  accept=".docx"
                  helpText="Necessario solo per generare file .docx."
                />
              </div>
              <div className="space-y-4">
                <h3 className="text-lg font-semibold">Scegli il Firmatario</h3>
                <RadioGroup value={signer} onValueChange={setSigner} className="space-y-2 pt-2">
                  <div className="flex items-center space-x-2"><RadioGroupItem value="Il Direttore del Corso" id="r1" /><Label htmlFor="r1">Il Direttore del Corso</Label></div>
                  <div className="flex items-center space-x-2"><RadioGroupItem value="Il Comandante del Centro" id="r2" /><Label htmlFor="r2">Il Comandante del Centro</Label></div>
                </RadioGroup>
                <p className="text-sm text-muted-foreground pt-2">Il titolo apparirà sopra: <strong>Col. Massimiliano Fortino</strong>.</p>
              </div>
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle>Passo 3: Anteprima e Generazione</CardTitle>
            <CardDescription>Controlla i dati estratti e genera i documenti finali.</CardDescription>
          </CardHeader>
          <CardContent className="space-y-6">
            {courseInfo && (
              <Card>
                <CardHeader><CardTitle className="flex items-center gap-2 text-base"><Info size={16} /> Dati del Corso</CardTitle></CardHeader>
                <CardContent className="text-sm space-y-1">
                  <p><strong>Titolo:</strong> {courseInfo.title}</p>
                  <p><strong>Sede:</strong> {courseInfo.location}</p>
                  <p><strong>Periodo:</strong> {courseInfo.period}</p>
                </CardContent>
              </Card>
            )}
            {discenti.length > 0 ? (
              <div className="space-y-2">
                <p><strong>Discenti Trovati:</strong> <span className="font-mono p-1 bg-muted rounded-md">{discenti.length}</span></p>
                <div className="rounded-md border max-h-[400px] overflow-y-auto">
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead>Matricola</TableHead>
                        <TableHead>Grado</TableHead>
                        <TableHead>Cognome</TableHead>
                        <TableHead>Nome</TableHead>
                        <TableHead>Categoria</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {discenti.map((d, index) => (
                        <TableRow key={index}>
                          <TableCell>{d.Matricola}</TableCell>
                          <TableCell>{d["Grado militare"]}</TableCell>
                          <TableCell>{d.Cognome}</TableCell>
                          <TableCell>{d.Nome}</TableCell>
                          <TableCell>{d.Categoria}</TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>
              </div>
            ) : (
              <div className="flex items-center justify-center text-center text-muted-foreground bg-muted/50 rounded-md p-8">
                <p>I dati dei discenti appariranno qui dopo il caricamento del file Excel.</p>
              </div>
            )}
            <div className="pt-4 flex flex-col space-y-4">
              <div className="flex items-center justify-center space-x-2">
                <Switch
                  id="output-format-switch"
                  checked={outputFormat === 'pdf'}
                  onCheckedChange={(checked) => setOutputFormat(checked ? 'pdf' : 'docx')}
                />
                <Label htmlFor="output-format-switch">Genera come PDF (raccomandato, qualità alta)</Label>
              </div>
              <Button size="lg" onClick={handleGenerateDocument} className="w-full" disabled={(!wordFile && outputFormat === 'docx') || discenti.length === 0}>
                <Download className="mr-2 h-5 w-5" /> Genera Modelli L
              </Button>
            </div>
          </CardContent>
        </Card>
        <PrintManager />
      </div>
    </div>
  );
}

export default Dashboard;