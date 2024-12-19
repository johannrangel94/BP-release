import * as React from 'react';
import { PrimaryButton } from '@fluentui/react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

interface XlsxExportButtonProps {
  items: Record<string, any>[]; // The data to export
  orderedItems: any[]; // The headers of the table
}

const XlsxExportButton: React.FC<XlsxExportButtonProps> = ({ items, orderedItems }) => {


  const exportToXLSX = () => {

    if (!items || items.length === 0) {
      return;
    }

    const headers = orderedItems.map((field: { text: any; }) => field.text);

    const rows = items.map(item => {
      return orderedItems.map((field: { key: string | number; }) => item[field.key] ?? ''); // Usa el valor de cada campo o '' si es undefined
    });

    const data = [headers, ...rows]; // La primera fila será la cabecera
  
    // 🔥 Usamos XLSX para convertir la matriz de datos a una hoja de Excel
    const worksheet = XLSX.utils.aoa_to_sheet(data); // aoa = Array of Arrays
  
    // 🔥 Creamos un libro de Excel y le añadimos la hoja de cálculo
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
  
    // 🔥 Exportamos el archivo Excel
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    
    // Usamos FileSaver para guardar el archivo
    const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    saveAs(blob, 'export.xlsx');

  }



  return (
    <PrimaryButton text="Export to XLSX" onClick={exportToXLSX} />
  );
};

export default XlsxExportButton;
