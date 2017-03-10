'use strict'
import * as pnp from 'sp-pnp-js'
export interface ISearchResult {
    link : string;
    title : string;
    description : string;
}

export interface ISearchService{
  GetSearchResults(query:string,resultsourceid:string,rowcount:string,pagenumber:number) : Promise<ISearchResult[]>;
}

export class MockSearchService implements ISearchService
{
    public GetSearchResults(query:string,resultsourceid:string,rowcount:string,pagenumber:number) : Promise<ISearchResult[]>{
        return new Promise<ISearchResult[]>((resolve,reject) => {
                resolve([
                    {title:'Title 1',description:'Title 1 desc',link:'http://asdada'},
                    {title:'Title 2',description:'Title 2 desc',link:'http://asdada'},
                    ]);
        });
    }
}

export class SearchService implements ISearchService
{
    public GetSearchResults(query:string,resultsourceid:string,rowcount:string,pagenumber:number) : Promise<ISearchResult[]>{
        let _results:ISearchResult[] = [];

        return new Promise<ISearchResult[]>((resolve,reject) => {
                pnp.sp.search('sharepoint')
                .then((results) => {
                   results.PrimarySearchResults.forEach((result)=>{
                    _results.push({
                        title:result.Title,
                        description:result.HitHighlightedSummary,
                        link:result.Path
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

