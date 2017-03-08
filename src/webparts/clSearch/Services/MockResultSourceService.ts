import * as pnp from "sp-pnp-js";

export interface IListService{
    getLists():Promise<string[]>;
}

export class MockListService implements IListService{

    private static _mockList:string[] = ['List 1','List 2', 'List 3'];

    public getLists():Promise<string[]>{
        return new Promise<string[]>((resolve) => {
           resolve(MockListService._mockList);
        });
    };
}

export class ListService implements IListService{
   
    public getLists():Promise<string[]>{
        debugger;
        let listsArray:string[] = [];
        return new Promise<string[]>((resolve) =>{

         pnp.sp.web.lists.get().then((lists) => {
            lists.forEach((list) => {listsArray.push(list.Title)});
              resolve(listsArray);
         });
          
        });
    }
}