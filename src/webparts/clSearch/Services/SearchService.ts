'use strict';
import * as pnp from 'sp-pnp-js';

export interface ISearchResult {
    link : string;
    title : string;
    description : string;
    author:string;
}

export interface ISearchService{
  GetSearchResults(query:string,resultsourceid:string,rowlimit:number,startrow:number) : Promise<ISearchResult[]>;
}

export class MockSearchService implements ISearchService
{
    public GetSearchResults(query:string,resultsourceid:string,rowlimit:number,startrow:number) : Promise<ISearchResult[]>{
        return new Promise<ISearchResult[]>((resolve,reject) => {
                resolve([
                    {title:'Title 1',description:'Title 1 desc',link:'http://asdada',author:'Pal'},
                    {title:'Title 2',description:'Title 2 desc',link:'http://asdada',author:'Pal'},
                    ]);
        });
    }
}

export class SearchService implements ISearchService
{
    public GetSearchResults(query:string,resultsourceid:string,rowlimit:number,startrow:number) : Promise<ISearchResult[]>{
        const _results:ISearchResult[] = [];

        return new Promise<ISearchResult[]>((resolve,reject) => {
                pnp.sp.search({
                     Querytext:query,
                     RowLimit:rowlimit,
                     StartRow:startrow
                    })
                .then((results) => {
                    console.log( results.PrimarySearchResults);
                   results.PrimarySearchResults.forEach((result)=>{
                    _results.push({
                        title:result.Title,
                        description:result.HitHighlightedSummary,
                        link:result.Path,
                        author:result.Author
                    });
                   });
                })
                .then(
                   () => { resolve(_results);}
                )
                .catch(
                    () => {reject(new Error("Some Error")); }
                );
                
        });
    }
}

