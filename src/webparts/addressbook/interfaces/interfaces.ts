import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Contact } from "../CRUD Operations/CRUD";

interface DialogData {
    handleCancelAdd: Function;
    handleCancelUpdate: Function;
    isUpdate?: boolean;
    contact?: Contact;
    context: WebPartContext;
    update: Function;
}

export interface ContactDataState {
    name: string,
    email: string,
    mobile: string,
    landline: string,
    website: string,
    address: string,
    nameWarning: string,
    mobileWarning: string,
    emailWarning: string
}

export interface ContactCRUDProps {
    context: WebPartContext
}

export default DialogData;