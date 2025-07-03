import React, { useState } from 'react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Card, CardHeader, CardTitle, CardContent } from '@/components/ui/card';
import { Table, TableHeader, TableRow, TableHead, TableBody, TableCell } from '@/components/ui/table';
import { showSuccess, showError } from '@/utils/toast';
import * as XLSX from 'xlsx';

interface Discente {
  "Matricola": string;
  "Grado militare": string;
  "Cognome e Nome del Discente": string;
  "Cognome": string;
  "Nome": string;
  "Categoria": string;
}

interface CourseInfo {
  "Denominazione del Corso": string;
  "Data inizio": string;
  "Data fine": string;
  "Sede": string;
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

        const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: "",
          blankrows: false,
        });

        // [Rest of your existing handleExcelUpload implementation...]
      } catch (error) {
        console.error("Error processing Excel file:", error);
        showError("Error processing Excel file");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div className="container mx-auto p-4">
      <Card className="mb-6">
        <CardHeader>
          <CardTitle>Upload Excel File</CardTitle>
        </CardHeader>
        <CardContent>
          <Input 
            type="file" 
            accept=".xlsx,.xls" 
            onChange={handleExcelUpload} 
            className="mb-4"
          />
          {excelFile && (
            <p className="text-sm text-gray-600">
              File selected: {excelFile.name}
            </p>
          )}
        </CardContent>
      </Card>

      {courseInfo && (
        <Card className="mb-6">
          <CardHeader>
            <CardTitle>Course Information</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <p className="font-medium">Course Name:</p>
                <p>{courseInfo["Denominazione del Corso"]}</p>
              </div>
              <div>
                <p className="font-medium">Location:</p>
                <p>{courseInfo["Sede"]}</p>
              </div>
              <div>
                <p className="font-medium">Start Date:</p>
                <p>{courseInfo["Data inizio"]}</p>
              </div>
              <div>
                <p className="font-medium">End Date:</p>
                <p>{courseInfo["Data fine"]}</p>
              </div>
            </div>
          </CardContent>
        </Card>
      )}

      {discenti.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle>Students List</CardTitle>
          </CardHeader>
          <CardContent>
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>ID</TableHead>
                  <TableHead>Rank</TableHead>
                  <TableHead>Last Name</TableHead>
                  <TableHead>First Name</TableHead>
                  <TableHead>Category</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {discenti.map((discente, index) => (
                  <TableRow key={index}>
                    <TableCell>{discente["Matricola"]}</TableCell>
                    <TableCell>{discente["Grado militare"]}</TableCell>
                    <TableCell>{discente["Cognome"]}</TableCell>
                    <TableCell>{discente["Nome"]}</TableCell>
                    <TableCell>{discente["Categoria"]}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
      )}
    </div>
  );
}

export default Dashboard;