/* eslint-disable @typescript-eslint/no-explicit-any */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";

//import { IItems } from "../../interfaces/IItems";
import { IFieldInfo } from "@pnp/sp/fields";
//import { IFieldInfo } from "@pnp/sp/fields";


export class SPService {

    private sp: SPFI;

    constructor(private context: WebPartContext) {
        this.sp = spfi().using(SPFx(this.context));
    }

    // Metodo para obtener todos los items de la lista que seleccionemos pero solo los campos que seleccionemos
    public getListItems = async (listId: string, selected: string[]): Promise<{items: Record<string, any>[], filteredFields: { title: string, internalName: string}[]}> => {
        try {
            // Obtiene todos los campos de la lista
            const fields = await this.getFields(listId);

            console.log("Todos los campos sin filtros de la lista son ", fields);
    
             //Filtramos los campos que tengan el mismo InternalName que los nombres seleccionados
             //en la Property Pane, esto por qué el PropertyFieldColumnPicker solo esta retornando los internalName
            const filteredFields = fields
              .filter(field => selected.includes(field.InternalName))
              .map(field => ({title: field.Title, internalName: field.InternalName})); // Filtramos los objetos completos de los campos seleccionados
              
    
            console.log("Desde service creamos objeto con internal y title name del field ", filteredFields); // Aquí verás los objetos completos de cada campo
    
            // Incluimos File_x0020_Type en la lista de campos a seleccionar
            const selectedFields = [...filteredFields.map(f => f.internalName), 'File_x0020_Type'].join(","); // Unimos todos los campos con una coma
            
            console.log("Campos seleccionados para la llamada de api son: ", selectedFields);
    
            // Consulta a la lista con los campos seleccionados
            const itemsRaw: any[] = await this.sp.web.lists.getById(listId).items.select(selectedFields).top(5)();

            console.log("Los items raw con File_x0020_Type a ver ", itemsRaw);
    
            // Creamos los objetos dinámicamente con los campos seleccionados
            // Tendra como clave el internalName de los fields seleccionados, y el valor sera el valor de su campo
            const items: Record<string, any>[] = itemsRaw.map((item) =>
                Object.fromEntries(
                    filteredFields.map(f => [(f.internalName === 'ContentType' ? 'File_x0020_Type' : f.internalName), item[(f.internalName === 'File_x0020_Type' ? 'File_x0020_Type' : f.internalName)]])
                )
            );
    
            console.log("Hemos obtenido los siguientes items con nombre interno del campo y su valor de item ", items);
    

            console.log("Procedemos a retornar los items y los campos que seleccionamos")
            return {items, filteredFields};
    
        } catch (error) {
            console.error(`Error obteniendo elementos de la lista ${listId} `, error);
            throw new Error(`No se pudieron obtener los elementos de la lista ${listId}: ${error.message}`);
        }
    }


        
    
    

    // Este metodo obtiene los campos de la lista que seleccionemos en el property pane, lo usamos para poder tener internalName y Title de todos los fields de la lista
    public getFields = async (listId: string): Promise<IFieldInfo[]> => {
        try {
            //const columns = await this.sp.web.lists.getById(listId).fields.filter("ReadOnlyField eq false and Hidden eq false")();
            const columns = await this.sp.web.lists.getById(listId).fields();
            return columns;
        } catch (error) {
            throw new error (`No se pudieron obtener los campos de la lista ${listId}`);

        }
    }


    


}