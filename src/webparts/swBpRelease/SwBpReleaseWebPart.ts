import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SwBpReleaseWebPartStrings';
import SwBpRelease from './components/SwBpRelease';
import { ISwBpReleaseProps } from './components/ISwBpReleaseProps';


import { PropertyFieldListPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';



export interface ISwBpReleaseWebPartProps {
  description: string;
  selectedList: string;
}

export default class SwBpReleaseWebPart extends BaseClientSideWebPart<ISwBpReleaseWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISwBpReleaseProps> = React.createElement(
      SwBpRelease,
      {
        context: this.context,
        description: this.properties.description,
        selectedList: this.properties.selectedList || ""
      }
    );

    ReactDom.render(element, this.domElement);
  }

  /*protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }*/


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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldListPicker('selectedList', {
                  label: "Select a list",
                  selectedList: this.properties.selectedList,
                  includeHidden: false,
                  disabled: false,
                  context: this.context,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'listPickerFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }



  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'list' && newValue !== oldValue) {
      this.properties.selectedList = newValue;
      this.render(); // Renderiza el componente nuevamente
    }

    // Importante llamar a este método para notificar el cambio al Property Pane
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }
}
