import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { IDropdownOption } from 'office-ui-fabric-react';
import { keys } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp"
export class SPOperations {

    public getAllListPNP(): Promise<IDropdownOption[]> {
        let listTitles: IDropdownOption[] = [];
        return new Promise<IDropdownOption[]>(async (resolve, reject) => {
            sp.web.lists.select('Title').get().then((results: any) => {
                results.map((result: any) => {
                    listTitles.push({
                        key: result.Title,
                        text: result.Title
                    })
                })
                resolve(listTitles);
            }, (error: any) => {
                reject("Error occured" + error)
            })
        })

    }

    public createListItePNP(listTitle: string): Promise<string> {

        return new Promise<string>(async (resolve, reject) => {
            sp.web.lists.getByTitle(listTitle).items.add({ Title: 'PNP ListItem' }).then((results: any) => {
                resolve("Item Added " + results.data.ID + " Successfully");
            }, (error: any) => {
                reject("Error Occured" + error)
            })
        })

    }


    public updateListItePNP(listTitle: string): Promise<string> {

        return new Promise<string>(async (resolve, reject) => {
            this.getLatestItemIdPNP(listTitle).then((itemid:number)=>{
                sp.web.lists.getByTitle(listTitle).items.getById(itemid).update({
                    Title:"PNPJS Update Item"
                }).then((result:any)=>{
                    resolve("Item Id " + itemid +" update")
                },(error:any)=>{
                    reject("Error occured" + error)
                })
            })
         })

    }
    
    public deleteListItePNP(listTitle: string): Promise<string> {

        return new Promise<string>(async (resolve, reject) => {
           this.getLatestItemIdPNP(listTitle).then((itemid:number)=>{
               sp.web.lists.getByTitle(listTitle).items.getById(itemid).delete().then((result:any)=>{
                   resolve("Item Id " + itemid +" deleted")
               },(error:any)=>{
                   reject("Error occured" + error)
               })
           })
        })

    }

    public getLatestItemIdPNP(listTitle: string): Promise<number> {
        return new Promise<number>(async (resolve, reject) => {
           sp.web.lists.getByTitle(listTitle).items.select('ID').orderBy('ID',false).top(1).get().then((result:any)=>{
               resolve(result[0].ID)
           },(error:any)=>{
               reject("Error Occoured" + error)
           })
        })

    }

    public getAllList(context: WebPartContext): Promise<IDropdownOption[]> {
        let restApiUrl: string = context.pageContext.web.absoluteUrl + '/_api/web/lists?select=Title';
        var listTitles: IDropdownOption[] = [];
        return new Promise<IDropdownOption[]>(async (resolve, reject) => {



            context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
                response.json().then((results: any) => {
                    console.log(results);
                    results.value.map((result: any) => {
                        listTitles.push({
                            key: result.Title,
                            text: result.Title
                        })
                    })
                });
                resolve(listTitles);
            }, (error: any): void => {
                reject('error occured' + error);
            });
        });
    }

    public createListTitle(context: WebPartContext, listTitle: string): Promise<string> {
        let restApiUrl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listTitle + "')/items";

        const body: string = JSON.stringify({
            Title: 'New Item Created'
        })

        const options: IHttpClientOptions = {
            headers: {
                Accept: 'application/json; odata=nometadata',
                "content-type": 'application/json; odata=nometadata',
                "OData-Version": ''
            },
            body: body
        }
        return new Promise<string>(async (resolve, reject) => {
            context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options).then((response: SPHttpClientResponse) => {
                response.json().then((result: any) => {
                    console.log(result);
                    resolve('Item with ID ' + result.ID + ' created successfully');
                }, (error: any) => {
                    reject('Error Occured ' + error)
                })
            })
        })
    }

    public deleteListItem(context: WebPartContext, listTitle: string): Promise<string> {

        let restApiUrl = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listTitle + "')/items";

        return new Promise<string>(async (resolve, reject) => {
            this.getLatestItemId(context, listTitle).then((itemId: number) => {
                context.spHttpClient.post(restApiUrl + "(" + itemId + ")", SPHttpClient.configurations.v1, { headers: { Accept: 'application/json; odata=nometadata', "content-type": 'application/json; odata=nometadata', "OData-Version": '', "IF-MATCH": "*", "X-HTTP-METHOD": "DELETE" } }).then((response: SPHttpClientResponse) => {
                    resolve("Item id " + itemId + " deleted successfully");
                }, (error: any) => {
                    reject("Error occured" + error);
                })
            })
        })

    }

    public getLatestItemId(context: WebPartContext, listTitle: string): Promise<number> {
        let restApiUrl = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listTitle + "')/items/?$orderBy=Id desc&$top=1&$select=id"

        return new Promise<number>(async (resolve, reject) => {
            context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json; odata=nometadata', "content-type": 'application/json; odata=nometadata', "OData-Version": '' } }).then((response: SPHttpClientResponse) => {
                response.json().then((result: any) => {
                    console.log(result);
                    resolve(result.value[0].Id);
                }, (error: any) => {
                    reject('Error occoured' + error)
                })
            })
        })
    }

    public updateListItem(context: WebPartContext, listTitle: string): Promise<string> {

        let restApiUrl = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listTitle + "')/items(22)";
        const body: string = JSON.stringify({
            Title: "Updated Item"
        })
        return new Promise<string>(async (resolve, reject) => {
            context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json; odata=nometadata', "content-type": 'application/json; odata=nometadata', "OData-Version": '', "IF-MATCH": "*", "X-HTTP-METHOD": "MERGE" }, body: body }).then((response: SPHttpClientResponse) => {
                resolve("Item id updated successfully");
            }, (error: any) => {
                reject("Error occured" + error);
            })
        })

    }
}