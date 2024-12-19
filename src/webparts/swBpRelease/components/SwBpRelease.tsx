/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
//import styles from './SwBpRelease.module.scss';
import type { ISwBpReleaseProps } from './ISwBpReleaseProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPService } from '../../SPService/SPService';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react';


import { FieldPicker } from "@pnp/spfx-controls-react/lib/FieldPicker";
import { FieldsOrderBy } from '@pnp/spfx-controls-react/lib/services/ISPService';
import { ISPField } from '@pnp/spfx-controls-react';

//import * as XLSX from 'xlsx';
//import { saveAs } from 'file-saver';
import XlsxExportButton from './xlsxExportButton/XlsxExportButton';
import CsvExportButton from './csvExportButton/CsvExportButton'


interface ISwBpReleaseState {
  items: Record<string,any>[]
  selectedFields: string[];
  filteredFields: { title: string, internalName: string }[];
  sortColumn?: string; // Nombre de la columna actual ordenada
  isSortDescending?: boolean; // Indica si la ordenación es descendente
  absoluteUrl: URL;
}

export default class SwBpRelease extends React.Component<ISwBpReleaseProps, ISwBpReleaseState> {

  private service: SPService;

  constructor(props: ISwBpReleaseProps){
    super(props);

    // Inicializamos estado de items vacio
    this.state = {
      items: [],
      selectedFields:[],
      filteredFields: [],
      absoluteUrl: new URL(this.props.context.pageContext.web.absoluteUrl)
    };
  }

  async componentDidMount(): Promise<void> {

      // Urls del sitio actual
      //const absoluteUrl = this.props.context.pageContext.web.absoluteUrl;
      //const originUrl = new URL(absoluteUrl).origin;

      // Obtenemos los items de los fields seleccionados en el propertypane


      // Items mapeados para representar en la lista
      /*const items: IItems[] = itemsRaw.map((item, index) => ({
        // Agregamos un id internamente si lo requieres
        id: index + 1,
        type: `${absoluteUrl}/_layouts/15/images/ic${item.type}.png`,
        name: item.name,
        technicalFunction: item.technicalFunction,
        product: item.product,
        oem: item.oem,
        Securitylevel: item.Securitylevel,
        url: `${originUrl}${item.url}`
      }));*/

      // Aqui estamos creando Title como clave y valor el item

      void (this.loadItems());

    }

    componentDidUpdate(prevProps: Readonly<ISwBpReleaseProps>, prevState: Readonly<ISwBpReleaseState>, snapshot?: any): void {
      if (prevProps.selectedList !== this.props.selectedList) {
        console.log("Se actualizo la lista a previsualizar");
        this.render();
      }

      if (prevProps.selectedFields !== this.props.selectedFields){
        console.log("Se ha actualizado los selected fields");
        void (this.loadItems());
        this.render();
      }


    }


    loadItems = async (): Promise<void> => {
      try {
        void (this.service = new SPService(this.props.context));

        const { items, filteredFields } = await this.service.getListItems(this.props.selectedList, this.props.selectedFields);

        console.log("Items desde SharePoint: ", items);
        console.log("Campos filtrados (Title ↔️ InternalName): ", filteredFields);
  
        this.setState({ items, filteredFields });

      } catch (error) {
        console.error("Error cargando los elementos ", error);
      }
    }

    onFieldPickerChanged = (fields: ISPField | ISPField[]): void => {
    // Verificamos si 'fields' es un array o un solo campo y lo convertimos a array
    const selectedFields = Array.isArray(fields) ? fields : [fields];
    
    // Extraemos los nombres internos (InternalName) de los campos seleccionados
    const fieldInternalNames = selectedFields.map(field => field.InternalName!);
    
    // Guardar los campos seleccionados en el estado del componente
    this.setState({ selectedFields: fieldInternalNames });

    console.log(`El field seleccionado desde pantalla es ${fieldInternalNames} y el seleccionado desde el panel de propiedades es ${this.props.selectedFields}`);
    }
    




  // Las columnas de la tabla, aqui podemos configurarlas
  /*private columns: IColumn[] = [
    { 
      key: 'type', 
      name: 'Type', 
      iconName: 'Document', 
      isIconOnly: true,  
      fieldName: 'type', 
      minWidth: 16, maxWidth: 16, 
      isResizable: true, 
      onRender: (item: IItems) => (
        <a href={item.url} download={true}>
          <img src={item.type} alt="File Type" style={{width: 16, height: 16}}/>
        </a>
      ) 
    },
    { key: 'name', name: 'Name', fieldName: 'name', minWidth: 100, isResizable: true },
    { key: 'technicalFunction', name: 'Technical Function', fieldName: 'technicalFunction', minWidth: 50, maxWidth: 50, isResizable: true },
    { key: 'product', name: 'Product', fieldName: 'product', minWidth: 100, isResizable: true },
    { key: 'oem', name: 'OEM', fieldName: 'oem', minWidth: 50, maxWidth: 50, isResizable: true },
    { 
      key: 'Securitylevel', 
      name: 'Security Level', 
      fieldName: 'Securitylevel', 
      minWidth: 250, 
      isResizable: true, 
      onRender: (item: IItems) => {
        let backgroundColor = '#ffffff';
        let border = '1px solid';
        let borderColor = '#ffffff';
        let color = '#ffffff';
        if(item.Securitylevel.startsWith("Level 1")){
          backgroundColor = '#BAEFCD';
          borderColor = '#76B16F';
          color = '#129103';
        } else if (item.Securitylevel.startsWith("Level 2")){
          backgroundColor = '#BBEFE9';
          borderColor = '#4FA155';
          color = '#4FA155';
        } else if (item.Securitylevel.startsWith("Level 3")){
          backgroundColor = '#FFECBF';
          borderColor = '#D55A5E';
          color = '##D55A5E';
        } else if (item.Securitylevel.startsWith("Level 4")) {
          backgroundColor = '#FFBCC5';
          borderColor = '#CC3939';
          color = '#CC3939';
        }

        return (
          <div style={{
            backgroundColor,
            border,
            borderColor,
            color,
            padding: '8px'
          }}>
            {item.Securitylevel}
          </div>
        );
      } 
    },
  ];*/


  private onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { items, sortColumn, isSortDescending } = this.state;
    const isSortedDescending = sortColumn === column.key ? !isSortDescending : false;
  
    const sortedItems = items.slice().sort((a, b) => {
      if (a[column.fieldName!] < b[column.fieldName!]) {
        return isSortedDescending ? 1 : -1;
      }
      if (a[column.fieldName!] > b[column.fieldName!]) {
        return isSortedDescending ? -1 : 1;
      }
      return 0;
    });
  
    this.setState({
      items: sortedItems,
      sortColumn: column.key,
      isSortDescending: isSortedDescending
    });
  }
  
  
/*
  generateColumns = (): IColumn[] => {

    if (!this.state.filteredFields || this.state.filteredFields.length === 0) return [];

    // SelectedFields retorna el internal name mientras que items esta usando Title
    const columns: IColumn[] = this.state.filteredFields.map(({title, internalName}) => ({
      key: internalName,
      name: title,
      fieldName: internalName,
      minWidth: 100,
      isResizable: true,
      isSorted: this.state.sortColumn === internalName, // Muestra que esta columna está ordenada
      isSortedDescending: this.state.sortColumn === internalName ? this.state.isSortDescending : false, // Indica si está en orden descendente
      onColumnClick: this.onColumnClick, // Llama a la función de ordenación
      onRender: (item: any) => {
        if (internalName === 'File_x0020_Type' && item[internalName]) {
          return (
            <img src={`${item.type}`} alt="File Type" style={{ width: 16, height: 16 }} />
          );
        }
        return <span>{item[internalName]}</span>;
      }
    })); 


    const columns: IColumn[] = this.props.orderedItems.map(({ key }) => {
      const field = this.state.filteredFields.find(f => f.internalName === key);
      if (!field) return null;

      return {
        key: field.internalName,
        name: field.title,
        fieldName: field.internalName,
        minWidth: 100,
        isResizable: true,
        isSorted: this.state.sortColumn === field.internalName,
        isSortedDescending: this.state.sortColumn === field.internalName ? this.state.isSortDescending : false,
        onColumnClick: this.onColumnClick,
        onRender: (item: any) => <span>{item[field.internalName]}</span>
      };
    }).filter(Boolean) as IColumn[];


    return columns;
  }
*/




    generateColumns = (): IColumn[] => {

      if (!this.state.filteredFields || this.state.filteredFields.length === 0) return [];
    
      const columns: IColumn[] = this.props.orderedItems.map(({ key }) => {
        const field = this.state.filteredFields.find(f => f.internalName === key);
        if (!field) return null;
    
        // 🔥 Comprobar si la columna es 'SecurityLevel' para aplicar la lógica especial
        if (field.internalName === 'Securitylevel') {
          return {
            key: field.internalName,
            name: field.title,
            fieldName: field.internalName,
            minWidth: 250, // 🔥 Esta columna tiene un ancho mínimo especial
            isRowHeader: true,
            isResizable: true,
            isSorted: this.state.sortColumn === field.internalName,
            isSortedDescending: this.state.sortColumn === field.internalName ? this.state.isSortDescending : false,
            onColumnClick: this.onColumnClick,
            onRender: (item: any) => {
              let backgroundColor = '#ffffff';
              let border = '1px solid';
              let borderColor = '#ffffff';
              let color = '#ffffff';
    
              // 🔥 Aplicar la lógica de colores dependiendo del valor de SecurityLevel
              if (item[field.internalName]?.startsWith("Level 1")) {
                backgroundColor = '#BAEFCD';
                borderColor = '#76B16F';
                color = '#129103';
              } else if (item[field.internalName]?.startsWith("Level 2")) {
                backgroundColor = '#BBEFE9';
                borderColor = '#4FA155';
                color = '#4FA155';
              } else if (item[field.internalName]?.startsWith("Level 3")) {
                backgroundColor = '#FFECBF';
                borderColor = '#D55A5E';
                color = '#D55A5E';
              } else if (item[field.internalName]?.startsWith("Level 4")) {
                backgroundColor = '#FFBCC5';
                borderColor = '#CC3939';
                color = '#CC3939';
              }
    
              return (
                <div style={{
                  backgroundColor,
                  border,
                  borderColor,
                  color,
                  padding: '8px'
                }}>
                  {item[field.internalName]}
                </div>
              );
            }
          };
        }

        if(field.internalName === 'DocIcon'){
          return { 
            key: 'type', 
            name: 'Type', 
            iconName: 'Document', 
            isIconOnly: true,  
            fieldName: 'type', 
            minWidth: 16, maxWidth: 16, 
            isResizable: true, 
            onRender: (item: any) => (
                <img src={`${this.state.absoluteUrl}/_layouts/15/images/ic${item[field.internalName]}.png`} alt="File Type" style={{width: 16, height: 16}}/>
            ) 
          }
            
          
        } 
    
        // 🔥 Lógica para otras columnas normales
        return {
          key: field.internalName,
          name: field.title,
          fieldName: field.internalName,
          minWidth: 100, 
          isResizable: true,
          isSorted: this.state.sortColumn === field.internalName, 
          isSortedDescending: this.state.sortColumn === field.internalName ? this.state.isSortDescending : false, 
          onColumnClick: this.onColumnClick, 
          onRender: (item: any) => <span>{item[field.internalName]}</span> 
        };
      }).filter(Boolean) as IColumn[]; // 🔥 Filtra los "null" para asegurar que solo se usen columnas válidas
    
      return columns;
    }
    



  /* Metodo para exportar a csv
  private exportToCSV = () => {
    const { items } = this.state;

    const { orderedItems } = this.props;

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
      ];
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
  }*/

/*
  private exportToXLSX = () => {
    const { items } = this.state;

    const { orderedItems } = this.props;

    if (!items || items.length === 0) {
      return;
    }

    const headers = orderedItems.map(field => field.text);

    const rows = items.map(item => {
      return orderedItems.map(field => item[field.key] ?? ''); // Usa el valor de cada campo o '' si es undefined
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
*/

  public render(): React.ReactElement<ISwBpReleaseProps> {
    

    const { orderedItems } = this.props;
    const { items } = this.state;


    const tableStyle = {
      root: {
        border: '1px solid #E5E5E5'
      }
    };

    return (
      <section>


    <FieldPicker
      context={this.props.context}
      includeHidden={false}
      includeReadOnly={false}
      label="Select your field(s)"
      multiSelect={true}
      orderBy={FieldsOrderBy.Title}
      listId={this.props.selectedList}
      onSelectionChanged={this.onFieldPickerChanged}
      showBlankOption={true}
    />



        {/* Agregamos un botón para exportar a CSV 
        <div style={{ marginBottom: '10px' , marginTop: '10px'}}>
          <PrimaryButton text="Exportar a CSV" onClick={this.exportToCSV} />
        </div>*/}

        <CsvExportButton items={items} orderedItems={orderedItems} />

        <XlsxExportButton items={items} orderedItems={orderedItems} />
        
        

        <DetailsList
          items={this.state.items}
          columns={this.generateColumns()}
          styles={tableStyle}
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          selectionMode={SelectionMode.none}
        />
      </section>
    );
  }
}
