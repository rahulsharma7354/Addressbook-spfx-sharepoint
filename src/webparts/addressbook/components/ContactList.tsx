import "./Addressbook.scss";
import * as React from 'react';
import { Component } from "react";
import { Contact } from "../CRUD Operations/CRUD";
import ContactCard from "./Cards";

class ContactCardList extends Component<{
  contacts: Array<Contact>;
  onClick: Function;
}> {
  constructor(props: any) {
    super(props);
    this.state = { contacts: this.props.contacts };
    this.showContact = this.showContact.bind(this);
  }

  showContact(id: string) {
    this.props.onClick(id);
  }

  render() {
    const contactList = this.props.contacts.map((contact) => {
      console.log(contact);
      return (
        <ContactCard
          key={contact.key}
          contact={contact}
          showContact={this.showContact}
        ></ContactCard>
      );
    });
    return <>{contactList}</>;
  }
}

export default ContactCardList;
