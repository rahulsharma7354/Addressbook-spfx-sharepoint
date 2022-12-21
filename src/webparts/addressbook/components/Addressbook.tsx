import * as React from "react";
import { Component } from "react";
import "./Addressbook.scss";
import Header from "./Header/Header";
import NavBar from "./Navbar/NavBar";
import { ApiProvider, Contact } from "../CRUD Operations/CRUD";
import ContactDetails from "./Body/PersonDetails";
import ContactCardList from "./Body/ContactList";
import AddContactDialog from "./Body/Contacts";
import DialogData, { ContactCRUDProps } from "../interfaces/interfaces";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export default class Addressbook extends Component<ContactCRUDProps, any> {
  constructor(props: ContactCRUDProps) {
    super(props);
    this.state = {
      showAddDialog: false,
      showUpdateDialog: false,
      savedContacts: [],
      selectedContact: null,
    };
    this.handleAddDialog = this.handleAddDialog.bind(this);
    this.handleUpdateDialog = this.handleUpdateDialog.bind(this);
    this.showContact = this.showContact.bind(this);
    this.setContacts = this.setContacts.bind(this);

    let apiProvider = new ApiProvider();
    apiProvider.checkExistence(this.props.context);
  }

  handleAddDialog() {
    this.setState({ showAddDialog: !this.state.showAddDialog });
  }

  handleUpdateDialog() {
    this.setState({ showUpdateDialog: !this.state.showUpdateDialog });
  }

  showContact(id: string) {
    let userList = this.state.savedContacts;
    function getData(x: Contact) {
      if (x.key.toString() === id) {
        return 1;
      }
    }
    let filterList = userList.filter(getData);
    let selected: Contact = filterList[0];
    this.setState({
      selectedContact: selected,
    });
  }

  componentDidMount() {
    this.setContacts(this.props.context);
  }

  setContacts(context: WebPartContext) {
    let contacts: Array<Contact> = [];
    ApiProvider.getContacts(context).then((data) => {
      data.value.map((item: any, index: number) => {
        let contact: Contact = {
          key: item.ID ?? "",
          name: item.Title ?? "",
          email: item.Email ?? "",
          mobile: item.Mobile ?? "",
          landline: item.Landline ?? "",
          website: item.Website ?? "",
          address: item.Address ?? "",
        };
        contacts.push(contact);
      });
      console.log(data.length);
      if (contacts.length !== 0) {
        this.setState({
          savedContacts: contacts,
          selectedContact: contacts[0],
        });
      }
    });
  }

  render() {
    let dialogDataUpdate: DialogData = {
      handleCancelAdd: this.handleAddDialog,
      handleCancelUpdate: this.handleUpdateDialog,
      isUpdate: true,
      contact: this.state.selectedContact,
      context: this.props.context,
      update: this.setContacts,
    };

    let dialogDataAdd: DialogData = {
      handleCancelAdd: this.handleAddDialog,
      handleCancelUpdate: this.handleUpdateDialog,
      context: this.props.context,
      update: this.setContacts,
    };

    return (
      <div className="scaffold">
        <Header></Header>
        <NavBar handleDialog={this.handleAddDialog}></NavBar>
        <div className="body">
          <div className="savedContacts">
            <p className="contactHeading">CONTACTS</p>
            {this.state.savedContacts.length !== 0 ? (
              <ContactCardList
                contacts={this.state.savedContacts}
                onClick={this.showContact}
              ></ContactCardList>
            ) : (
              <div className="emptyMessage">No Contacts</div>
            )}
          </div>
          {this.state.selectedContact !== null ? (
            <ContactDetails
              contact={this.state.selectedContact}
              handleDialog={this.handleUpdateDialog}
              context={this.props.context}
              setContacts={this.setContacts}
            ></ContactDetails>
          ) : null}
          {this.state.showUpdateDialog === true ? (
            <AddContactDialog dialogData={dialogDataUpdate}></AddContactDialog>
          ) : null}
          {this.state.showAddDialog === true ? (
            <AddContactDialog dialogData={dialogDataAdd}></AddContactDialog>
          ) : null}
        </div>
      </div>
    );
  }
}
