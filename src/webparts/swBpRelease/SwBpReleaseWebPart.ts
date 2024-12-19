import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SwBpReleaseWebPartStrings';
import SwBpRelease from './components/SwBpRelease';
import { ISwBpReleaseProps } from './components/ISwBpReleaseProps';


import { PropertyFieldListPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { IColumnReturnProperty, PropertyFieldColumnPicker, PropertyFieldColumnPickerOrderBy } from '@pnp/spfx-property-controls';
import { PropertyFieldOrder } from '@pnp/spfx-property-controls/lib/PropertyFieldOrder';

export interface ISwBpReleaseWebPartProps {
  selectedList: string;
  selectedFields: string[];
  orderedItems: Array<any>;
}

export default class SwBpReleaseWebPart extends BaseClientSideWebPart<ISwBpReleaseWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISwBpReleaseProps> = React.createElement(
      SwBpRelease,
      {
        context: this.context,
        selectedList: this.properties.selectedList || "",
        selectedFields: this.properties.selectedFields || [],
        orderedItems: this.properties.orderedItems || []
      }
    );
    ReactDom.render(element, this.domElement);


    console.log("Esto se ejecuta cuando se inicia el webpart, selectedFields tiene el valor ", this.properties.selectedFields, " ordered items tiene ", this.properties.orderedItems);


  }


  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Descripcion del webpart'
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker('selectedList', {
                  label: "Select a list",
                  selectedList: this.properties.selectedList,
                  includeHidden: false,
                  disabled: false,
                  context: this.context,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldColumnPicker('selectedFields', {
                  label: "Select columns",
                  context: this.context,
                  selectedColumn: this.properties.selectedFields,
                  listId: this.properties.selectedList,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'multiColumnPickerFieldId',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty['Internal Name'],
                  multiSelect: true//,
                  //filter: "ReadOnlyField eq false and Hidden eq false"
                }),
                PropertyFieldOrder("orderedItems", {
                  key: "orderedItems",
                  label: "Ordered Items",
                  textProperty: "text",
                  items: this.properties.orderedItems && this.properties.orderedItems.length > 0
                  ? (this.properties.orderedItems)
                  : (this.properties.selectedFields || []).map((field) => ({
                      key: field,
                      text: field
                    })),
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged
                })
                
              ]
            }
          ]
        }
      ]
    };
  }



  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  /*protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'selectedList' && newValue !== oldValue) {
      this.properties.selectedList = newValue;
      this.render(); // Renderiza el componente nuevamente
    }

    if (propertyPath === 'selectedFields' && newValue !== oldValue) {
      this.properties.selectedFields = newValue;
      // Solo actualizar orderedItems si no hay uno existente o está vacío
      if (!this.properties.orderedItems || this.properties.orderedItems.length === 0) {
        this.properties.orderedItems = newValue.map((field: any) => ({
          key: field,
          text: field
        }));
      }
      this.context.propertyPane.refresh();
      this.render();

      console.log("Se actualiazo selectedfields", this.properties.selectedFields);

    }
    

    if (propertyPath === 'orderedItems' && newValue !== oldValue) {
      this.properties.orderedItems = newValue;

      console.log("Se actualizo ordered items ", this.properties.orderedItems);

    }


    // Importante llamar a este método para notificar el cambio al Property Pane
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }*/

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
  if (propertyPath === 'selectedList' && newValue !== oldValue) {
    this.properties.selectedList = newValue;
    this.render(); // Renderiza el componente nuevamente
  }

  if (propertyPath === 'selectedFields' && newValue !== oldValue) {
    this.properties.selectedFields = newValue;

    // 🔥 Nueva lógica para sincronizar orderedItems con selectedFields
    const newOrderedItems = newValue.map((field: any) => ({
      key: field,
      text: field
    }));

    if (!this.properties.orderedItems || this.properties.orderedItems.length === 0) {
      // Si orderedItems no existe o está vacío, lo inicializamos con los nuevos campos
      this.properties.orderedItems = newOrderedItems;
    } else {
      // 🔥 Agregar nuevos campos que no están en orderedItems
      newOrderedItems.forEach((newItem: { key: any; }) => {
        const existingItem = this.properties.orderedItems.find((item) => item.key === newItem.key);
        if (!existingItem) {
          this.properties.orderedItems.push(newItem); // Añadir los nuevos campos seleccionados
        }
      });

      // 🔥 (Opcional) Eliminar los campos de orderedItems que ya no están en selectedFields
      this.properties.orderedItems = this.properties.orderedItems.filter((item) =>
        newValue.includes(item.key)
      );
    }

    console.log("🔄 Se actualizó orderedItems", this.properties.orderedItems);

    // Refresca el Property Pane para mostrar la nueva lista de orden
    this.context.propertyPane.refresh();

    // Renderizar el componente para actualizar la tabla
    this.render();

    console.log("⚙️ Se actualizó selectedFields", this.properties.selectedFields);
  }

  if (propertyPath === 'orderedItems' && newValue !== oldValue) {
    this.properties.orderedItems = newValue;
    console.log("🔄 Se actualizó orderedItems", this.properties.orderedItems);
  }

  // Importante llamar a este método para notificar el cambio al Property Pane
  super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
}


  }
