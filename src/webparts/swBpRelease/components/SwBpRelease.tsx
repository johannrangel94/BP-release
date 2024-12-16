/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
//import styles from './SwBpRelease.module.scss';
import type { ISwBpReleaseProps } from './ISwBpReleaseProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPService } from '../../SPService/SPService';
import { DetailsList, DetailsListLayoutMode, IColumn } from '@fluentui/react';
import { IItems } from '../../../interfaces/IItems';

interface ISwBpReleaseState {
  items: IItems[];
}

export default class SwBpRelease extends React.Component<ISwBpReleaseProps, ISwBpReleaseState> {

  private service: SPService;

  constructor(props: ISwBpReleaseProps){
    super(props);

    // Inicializamos estado de items vacio
    this.state = {
      items: []
    };
  }

  async componentDidMount(): Promise<void> {
      console.log("Se procedera a instanciar el service con el context");

      const absoluteUrl = this.props.context.pageContext.web.absoluteUrl;

      const originUrl = new URL(absoluteUrl).origin;


      this.service = new SPService(this.props.context);
      const itemsRaw = await this.service.getListItems(this.props.selectedList);
      //console.log(`Los items son ${JSON.stringify(this.items2, null, 2)}`);
      const items: IItems[] = itemsRaw.map((item, index) => ({
        id: index + 1,
        type: `${absoluteUrl}/_layouts/15/images/ic${item.type}.png`,
        name: item.name,
        technicalFunction: item.technicalFunction,
        product: item.product,
        oem: item.oem,
        securityLevel: item.securityLevel,
        url: `${originUrl}${item.url}`
      }));

      this.setState({ items });
      console.log(`Se instancio el estado de la clase: ${this.state.items}`);


    }


  // Las columnas de la tabla (solo 2 columnas)
  private columns: IColumn[] = [
    { key: 'type', name: 'Type', iconName: 'Document', isIconOnly: true,  fieldName: 'type', minWidth: 16, maxWidth: 16, isResizable: true, onRender: (item: IItems) => (
      <a href={item.url} download={true}><img src={item.type} alt="File Type" style={{width: 16, height: 16}}/></a>
    ) },
    { key: 'name', name: 'Name', fieldName: 'name', minWidth: 100, isResizable: true },
    { key: 'technicalFunction', name: 'Technical Function', fieldName: 'technicalFunction', minWidth: 50, maxWidth: 50, isResizable: true },
    { key: 'product', name: 'Product', fieldName: 'product', minWidth: 100, isResizable: true },
    { key: 'oem', name: 'OEM', fieldName: 'oem', minWidth: 50, maxWidth: 50, isResizable: true },
    { key: 'securityLevel', name: 'Security Level', fieldName: 'securityLevel', minWidth: 100, isResizable: true, onRender: (item: IItems) => {
      let backgroundColor = '#ffffff';
      if(item.securityLevel.startsWith("Level 1")){
        backgroundColor = '#BAEFCD'
      } else if (item.securityLevel.startsWith("Level 2")){
        backgroundColor = '#BBEFE9'
      } else if (item.securityLevel.startsWith("Level 3")){
        backgroundColor = '#FFECBF'
      } else if (item.securityLevel.startsWith("Level 4")) {
        backgroundColor = '#FFBCC5'
      }

      return (
        <div style={{
          backgroundColor,
          padding: '8px'
        }}>

          {item.securityLevel}

        </div>
      )

    } },
  ];
  



  public render(): React.ReactElement<ISwBpReleaseProps> {
    /*const {
      description,
      selectedList
    } = this.props;*/


    return (
      <section >

        <DetailsList
          items={this.state.items}
          columns={this.columns}
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
        />
      </section>
    );
  }
}
