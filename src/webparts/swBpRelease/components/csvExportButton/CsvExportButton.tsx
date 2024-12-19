import * as React from 'react';
import { PrimaryButton } from '@fluentui/react';


interface CsvExportButtonProps {
    items: Record<string, any>[];
    orderedItems: any[];
}


const CsvExportButton: React.FC<CsvExportButtonProps> = ({ items, orderedItems }) => {


      // Metodo para exportar a csv
      const exportToCSV = () => {
    
        if (!items || items.length === 0) {
          return;
        }
    
        // Definimos las columnas a exportar
        // Puedes ajustar el orden o las columnas que quieres exportar
        //const headers = ["name", "technicalFunction", "product", "oem", "Securitylevel", "url"];
    
        //const headers = this.state.filteredFields.map(f => f.title)
        
        const headers = orderedItems.map(field => field.text);
    
        // Crear las filas
        const rows = items.map((item) => {
    
          //return this.state.filteredFields.map(f => item[f.internalName]);
          return orderedItems.map( field => item[field.key]);
          /*return [
            item.name,
            item.technicalFunction,
            item.product,
            item.oem,
            item.Securitylevel,
            item.url
          ];*/
        });
    
        // Generar el contenido CSV
        let csvContent = headers.join(",") + "\n";
        rows.forEach(row => {
          // Escapar comas si es necesario, por simplicidad se asume que los datos no contienen comas
          csvContent += row.join(",") + "\n";
        });
    
        // Crear el Blob
        const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
    
        // Crear un link temporal y simular el click
        const link = document.createElement("a");
        link.href = url;
        link.setAttribute("download", "export.csv");
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      }


    return (
        <PrimaryButton text="Export to CSV" onClick={exportToCSV} />
    );
}

export default CsvExportButton;