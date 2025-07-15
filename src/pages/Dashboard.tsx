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

        let categoriaPerTemplate = discente.Categoria;
        const gradoMilitare = discente["Grado militare"].toUpperCase();

        // Updated mapping logic based on the provided image and new rules
        if (['GCA', 'GDV', 'GDB', 'COL', 'TCL', 'MAG', 'CAP', 'TEN', 'STN'].includes(gradoMilitare)) {
          categoriaPerTemplate = "l'Ufficiale";
        } else if (['LGT.CS', 'LGT', 'MAR.A', 'MAR.C', 'MAR.O', 'MAR'].includes(gradoMilitare)) {
          categoriaPerTemplate = "l'Ispettore";
        } else if (['B.C.QS', 'BRIG.C', 'BRIG', 'VBRIG'].includes(gradoMilitare)) {
          categoriaPerTemplate = "il Sovrintendente";
        } else if (['APS.QS', 'APP.SC', 'APP', 'FIN.SC', 'FIN'].includes(gradoMilitare)) {
          categoriaPerTemplate = "il Militare";
        } else {
          // Fallback if no specific match is found, use the category from Excel
          categoriaPerTemplate = discente.Categoria;
        }

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
        outputZip.file(`Attestato_${discente["Cognome e Nome del Discente"].replace(/[^a-zA-Z0-9]/g, '_')}.docx`, out);
      }
      const zipBlob = await outputZip.generateAsync({ type: "blob" });
      saveAs(zipBlob, "documenti_individuali.zip");
      dismissToast(toastId);
      showSuccess("Archivio ZIP con documenti DOCX generato con successo!");
    } catch (error: any) {
      dismissToast(toastId);
      console.error("Errore nella generazione del documento:", error);
      showError(`Errore durante la generazione: ${error.message}`);
    }
  };

  return (
    <div className="container mx-auto p-4 md:p-6 lg:p-8">
      <header className="text-center mb-8">
        <h1 className="text-3xl font-bold tracking-tight">
          Gestione Modelli L <span className="text-sm font-normal text-muted-foreground">v15.1.Lug25</span>
        </h1>
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
                  label="Template Documento"
                  file={wordFile}
                  onUpload={handleWordUpload}
                  onRemove={() => setWordFile(null)}
                  accept=".docx"
                  helpText="Carica il template Word (.docx) con i segnaposto."
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
            <div className="pt-4 flex flex-col space-y-2">
              <Button size="lg" onClick={handleGenerateDocument} className="w-full" disabled={!wordFile || discenti.length === 0}>
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