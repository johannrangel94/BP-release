/* eslint-disable @typescript-eslint/no-explicit-any */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { IItems } from "../../interfaces/IItems";


export class SPService {

    private sp: SPFI;


    constructor(private context: WebPartContext) {
        this.sp = spfi().using(SPFx(this.context));
    }


    public getListItems = async (listId: string): Promise<IItems[]> => {
        try {



            const itemsRaw = await this.sp.web.lists.getById(listId).items.select('FileLeafRef', 'File_x0020_Type', 'TechnicalFunction', 'Product', 'OEM', 'Securitylevel', 'FileRef').top(5)();

            const items: IItems[] = itemsRaw.map((item) => ({
                type: item.File_x0020_Type,
                name: item.FileLeafRef, 
                technicalFunction: item.TechnicalFunction,
                product: item.Product,
                oem: item.OEM,
                securityLevel: item.Securitylevel,
                url: item.FileRef
            }));


            return items;
        } catch (error) {
            console.error(`Error obteniendo elementos de la lista ${listId} `, error);
            throw new error (`No se pudieron obtener los elementos de la lista ${listId}: ${error.message}`);
        }
    }


    


}