import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export class ApiProvider {
    public checkListExistence(context: WebPartContext): void {
        const listUrl = context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('Contacts')";

        context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
            if (response.status == 200) {
                console.log("List Found");
                return;
            } else if (response.status === 404) {
                console.log("Not Found");
                return;
            }
        });

        console.log(listUrl);
    }

    static async getAllContacts(context: WebPartContext) {
        var requestInit = {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'odata-version': ''
            }
        };
        return new Promise<any>((resolve: (data: any) => void, reject: (error: any) => void): void => {
            context.spHttpClient.get(context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('Contacts')/items`, SPHttpClient.configurations.v1, requestInit).then((response: SPHttpClientResponse) => {
                return response.json().then((res) => {
                    if (!response.ok) {
                        reject({
                            status: response.status ? response.status.toString() : "",
                            message: (res.error.message && res.error.message.value) || 'Request Failed'
                        });
                    }
                    else {
                        resolve(res);
                    }
                });
            }).catch((e: any) => {
                reject({
                    status: "",
                    message: 'Request Failed'
                });
            });
        });
    }
    
    static createContact(contact: Contact, context: WebPartContext, update: Function): void {
        const body: string = JSON.stringify({
            // 'Key': contact.key,
            'Title': contact.name,
            'Email': contact.email,
            'Mobile': contact.mobile,
            'Landline': contact.landline,
            'Website': contact.website,
            'Address': contact.address
        });
        console.log(body);
        context.spHttpClient.post(`${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Contacts')/items`,
        SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: body
      })
            .then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    response.json().then((responseJSON) => {
                        console.log(responseJSON);
                        alert(`Contact created successfully with ID: ${responseJSON.ID}`);
                        update(context);
                    });
                } else {
                    response.json().then((responseJSON) => {
                        console.log(responseJSON);
                        alert(`Something went wrong! Check the error in the browser console.`);
                    });
                }
            }).catch(error => {
                console.log(error);
            });
    }

    static deleteContact(context: WebPartContext, key: string, update: Function) {
        context.spHttpClient.post(`${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Contacts')/items(${key})`,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-type': 'application/json;odata=verbose',
                    'odata-version': '',
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'DELETE'
                }
            })
            .then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    alert(`Contact deleted successfully!`);
                    update(context);
                }
                else {
                    alert(`Something went wrong!`);
                    console.log(response.json());
                }
            });
    }

    static updateContact(context: WebPartContext, contact: Contact, update: Function): void {
        const body: string = JSON.stringify({
            // 'Key': contact.key.toString(),
            'Title': contact.name,
            'Email': contact.email,
            'Mobile': contact.mobile,
            'Landline': contact.landline,
            'Website': contact.website,
            'Address': contact.address
        });
        console.log(`${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Contacts')/items(${contact.key})`);
        if (parseInt(contact.key) > 0) {
            context.spHttpClient.post(`${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Contacts')/items(${contact.key})`,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-type': 'application/json;odata=nometadata',
                        'odata-version': '',
                        'IF-MATCH': '*',
                        'X-HTTP-Method': 'MERGE'
                    },
                    body: body
                })
                .then((response: SPHttpClientResponse) => {
                    if (response.ok) {
                        alert(`Contact with ID: ${contact.key} updated successfully!`);
                        update(context);
                    } else {
                        response.json().then((responseJSON) => {
                            console.log(responseJSON);
                            alert(`Something went wrong! Check the error in the browser console.`);
                        });
                    }
                }).catch(error => {
                    console.log(error);
                });
        }
        else {
            alert(`Please select a valid contact`);
        }
    }
}

export interface Contact {
    key: string,
    name: string,
    email: string,
    mobile: string,
    landline: string,
    website: string,
    address: string
}
