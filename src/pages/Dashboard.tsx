import React, { useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Label } from "@/components/ui/label";
import { showSuccess, showError } from "@/utils/toast";

interface CourseInfo {
  courseName: string;
  courseCode: string;
  startDate: string;
  endDate: string;
  location: string;
  instructor: string;
}

interface Discente {
  Matricola: string;
  "Grado militare": string;
  "Cognome e Nome del Discente": string;
  Cognome: string;
  Nome: string;
  Categoria: string;
}

function Dashboard() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [discenti, setDiscenti] = useState<Discente[]>([]);
  const [courseInfo, setCourseInfo] = useState<CourseInfo | null>(null);

  const handleExcelUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setExcelFile(file);
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        setDiscenti([]);
        setCourseInfo(null);

        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Estrazione dati corso dalla riga 7 (indice 6)
        const courseNameCell = worksheet["C7"]?.v;
        const courseCodeCell = worksheet["C8"]?.v;
        const startDateCell = worksheet["C9"]?.v;
        const endDateCell = worksheet["C10"]?.v;
        const locationCell = worksheet["C11"]?.v;
        const instructorCell = worksheet["C12"]?.v;

        if (courseNameCell && courseCodeCell && startDateCell && endDateCell && locationCell && instructorCell) {
          setCourseInfo({
            courseName: String(courseNameCell),
            courseCode: String(courseCodeCell),
            startDate: new Date(startDateCell).toLocaleDateString("it-IT"),
            endDate: new Date(endDateCell).toLocaleDateString("it-IT"),
            location: String(locationCell),
            instructor: String(instructorCell),
          });
        } else {
          showError("Impossibile estrarre tutte le informazioni del corso dalla riga 7-12. Controlla il formato del file.");
        }

        const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: "",
          blankrows: false,
        });

        // Trova dinamicamente la riga delle intestazioni
        let headerRowIndex = -1;
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i].map(cell => String(cell).toLowerCase().trim());
            const hasMatricola = row.some(cell => cell.includes('matricola'));
            const hasCognome = row.some(cell => cell.includes('cognome'));
            const hasNome = row.some(cell => cell.includes('nome'));
            const hasGrado = row.some(cell => cell.includes('grado'));

            if (hasMatricola && (hasCognome || hasGrado)) {
                headerRowIndex = i;
                break;
            }
        }

        if (headerRowIndex === -1) {
            showError("Impossibile trovare la riga delle intestazioni. Assicurati che il file Excel contenga colonne come 'matricola', 'grado' e 'cognome'.");
            return;
        }

        const headers = jsonData[headerRowIndex].map(h => String(h).toLowerCase().trim());
        const dataRows = jsonData.slice(headerRowIndex + 1);

        if (dataRows.length === 0) {
          showError("Nessun discente trovato dopo la riga delle intestazioni. Controlla che il file Excel contenga dati validi.");
          return;
        }

        const findIndex = (keywords: string[]) => headers.findIndex(h => keywords.some(kw => h.includes(kw)));

        const matricolaIndex = findIndex(['matricola']);
        const gradoIndex = findIndex(['grado']);
        const categoriaIndex = findIndex(['cat.', 'cat', 'categoria']);
        
        const cognomeNomeIndex = findIndex(['cognome e nome', 'nominativo']);
        const cognomeIndex = findIndex(['cognome']);
        const nomeIndex = findIndex(['nome']);

        const valueExists = (val: any) => val !== null && val !== undefined && String(val).trim() !== '';

        const discentiData = dataRows.map((row, rowIndex) => {
          if (row.every(cell => !valueExists(cell))) {
            return null; // Salta righe completamente vuote
          }

          let cognomeNome: string = '';
          let cognome: string = '';
          let nome: string = '';

          // Se ci sono colonne separate per Cognome e Nome, usiamo quelle
          if (cognomeIndex !== -1 && nomeIndex !== -1) {
              cognome = valueExists(row[cognomeIndex]) ? String(row[cognomeIndex]) : '';
              nome = valueExists(row[nomeIndex]) ? String(row[nomeIndex]) : '';
              cognomeNome = `${cognome} ${nome}`.trim();
          } 
          // Altrimenti se c'è una colonna unica "Cognome e Nome", la dividiamo
          else if (cognomeNomeIndex !== -1 && valueExists(row[cognomeNomeIndex])) {
              const fullName = String(row[cognomeNomeIndex]).trim();
              const lastSpaceIndex = fullName.lastIndexOf(' ');
              
              if (lastSpaceIndex > 0) {
                  cognome = fullName.substring(0, lastSpaceIndex);
                  nome = fullName.substring(lastSpaceIndex + 1);
              } else {
                  cognome = fullName;
                  nome = '';
              }
              cognomeNome = fullName;
          }

          const matricola = matricolaIndex !== -1 ? row[matricolaIndex] : undefined;
          const grado = gradoIndex !== -1 ? row[gradoIndex] : undefined;
          const categoria = categoriaIndex !== -1 && valueExists(row[categoriaIndex]) ? row[categoriaIndex] : "";

          if (!valueExists(matricola) || !valueExists(grado) || !valueExists(cognomeNome)) {
            console.warn(`Riga ${headerRowIndex + rowIndex + 2} del file Excel saltata perché mancano dati essenziali (Matricola, Grado o Cognome/Nome). Dati letti:`, { matricola, grado, cognomeNome });
            return null;
          }

          return {
            "Matricola": String(matricola),
            "Grado militare": String(grado),
            "Cognome e Nome del Discente": cognomeNome,
            "Cognome": cognome,
            "Nome": nome,
            "Categoria": String(categoria),
          };
        }).filter(d => d !== null) as Discente[];

        if (discentiData.length === 0) {
          showError("Nessun discente valido caricato. Controlla che le intestazioni e i dati siano corretti e completi.");
        } else {
          setDiscenti(discentiData);
          showSuccess(`Caricamento completato. Trovati ${discentiData.length} discenti e dati del corso.`);
        }
      } catch (error) {
        console.error("Errore imprevisto durante la lettura del file Excel:", error);
        showError("Errore durante l'elaborazione del file. Assicurati che sia un file .xlsx valido e non corrotto.");
        setDiscenti([]);
        setCourseInfo(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const generateDocx = async () => {
    if (!courseInfo || discenti.length === 0) {
      showError("Carica prima un file Excel con i dati del corso e dei discenti.");
      return;
    }

    try {
      // Esempio di template DOCX (dovresti caricare il tuo template reale)
      // Per semplicità, qui usiamo un template fittizio.
      // In un'applicazione reale, caricheresti un file .docx dal server o da un input utente.
      const templateContent = `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>Corso: {courseName} ({courseCode})</w:t></w:r></w:p>
            <w:p><w:r><w:t>Periodo: dal {startDate} al {endDate}</w:t></w:r></w:p>
            <w:p><w:r><w:t>Sede: {location}</w:t></w:r></w:p>
            <w:p><w:r><w:t>Istruttore: {instructor}</w:t></w:r></w:p>
            <w:p><w:r><w:t>Elenco Discenti:</w:t></w:r></w:p>
            <w:tbl>
              <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
              <w:tr>
                <w:tc><w:p><w:r><w:t>Matricola</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Grado</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Cognome e Nome</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Categoria</w:t></w:r></w:p></w:tc>
              </w:tr>
              {#discenti}
              <w:tr>
                <w:tc><w:p><w:r><w:t>{Matricola}</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>{Grado militare}</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>{Cognome e Nome del Discente}</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>{Categoria}</w:t></w:r></w:p></w:tc>
              </w:tr>
              {/discenti}
            </w:tbl>
          </w:body>
        </w:document>
      `;
      
      // Per un template reale, useresti fetch o un input file:
      // const response = await fetch('/path/to/your/template.docx');
      // const buf = await response.arrayBuffer();
      // const zip = new PizZip(buf);

      // Per questo esempio, creiamo un zip fittizio per il template
      const zip = new PizZip();
      zip.file("word/document.xml", templateContent); // Docxtemplater si aspetta questo percorso

      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      doc.setData({
        courseName: courseInfo.courseName,
        courseCode: courseInfo.courseCode,
        startDate: courseInfo.startDate,
        endDate: courseInfo.endDate,
        location: courseInfo.location,
        instructor: courseInfo.instructor,
        discenti: discenti,
      });

      doc.render();

      const out = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        compression: "DEFLATE",
      });

      saveAs(out, `Elenco_Discenti_${courseInfo.courseCode}.docx`);
      showSuccess("Documento DOCX generato con successo!");
    } catch (error) {
      console.error("Errore durante la generazione del DOCX:", error);
      showError("Errore durante la generazione del documento DOCX.");
    }
  };

  const generateExcel = () => {
    if (!courseInfo || discenti.length === 0) {
      showError("Carica prima un file Excel con i dati del corso e dei discenti.");
      return;
    }

    try {
      const ws_data = [
        ["Informazioni Corso"],
        ["Nome Corso:", courseInfo.courseName],
        ["Codice Corso:", courseInfo.courseCode],
        ["Data Inizio:", courseInfo.startDate],
        ["Data Fine:", courseInfo.endDate],
        ["Sede:", courseInfo.location],
        ["Istruttore:", courseInfo.instructor],
        [], // Riga vuota per separazione
        ["Elenco Discenti"],
        ["Matricola", "Grado militare", "Cognome e Nome del Discente", "Categoria"],
        ...discenti.map(d => [d.Matricola, d["Grado militare"], d["Cognome e Nome del Discente"], d.Categoria])
      ];

      const ws = XLSX.utils.aoa_to_sheet(ws_data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Discenti e Corso");
      XLSX.writeFile(wb, `Elenco_Discenti_${courseInfo.courseCode}.xlsx`);
      showSuccess("File Excel generato con successo!");
    } catch (error) {
      console.error("Errore durante la generazione del file Excel:", error);
      showError("Errore durante la generazione del file Excel.");
    }
  };

  return (
    <div className="container mx-auto p-4">
      <h1 className="text-3xl font-bold text-center mb-8">Gestione Discenti Corsi</h1>

      <Card className="mb-6">
        <CardHeader>
          <CardTitle>Carica File Excel</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="grid w-full max-w-sm items-center gap-1.5">
            <Label htmlFor="excel-upload">Seleziona il file Excel (.xlsx)</Label>
            <Input
              id="excel-upload"
              type="file"
              accept=".xlsx"
              onChange={handleExcelUpload}
              className="cursor-pointer"
            />
          </div>
        </CardContent>
      </Card>

      {courseInfo && (
        <Card className="mb-6">
          <CardHeader>
            <CardTitle>Informazioni Corso</CardTitle>
          </CardHeader>
          <CardContent className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <p><strong>Nome Corso:</strong> {courseInfo.courseName}</p>
            <p><strong>Codice Corso:</strong> {courseInfo.courseCode}</p>
            <p><strong>Data Inizio:</strong> {courseInfo.startDate}</p>
            <p><strong>Data Fine:</strong> {courseInfo.endDate}</p>
            <p><strong>Sede:</strong> {courseInfo.location}</p>
            <p><strong>Istruttore:</strong> {courseInfo.instructor}</p>
          </CardContent>
        </Card>
      )}

      {discenti.length > 0 && (
        <Card className="mb-6">
          <CardHeader>
            <CardTitle>Elenco Discenti ({discenti.length})</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="overflow-x-auto">
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
                  {discenti.map((discente, index) => (
                    <TableRow key={index}>
                      <TableCell>{discente.Matricola}</TableCell>
                      <TableCell>{discente["Grado militare"]}</TableCell>
                      <TableCell>{discente.Cognome}</TableCell>
                      <TableCell>{discente.Nome}</TableCell>
                      <TableCell>{discente.Categoria}</TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
            <div className="flex justify-end space-x-4 mt-6">
              <Button onClick={generateDocx}>Genera DOCX</Button>
              <Button onClick={generateExcel}>Genera Excel</Button>
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}

export default Dashboard;