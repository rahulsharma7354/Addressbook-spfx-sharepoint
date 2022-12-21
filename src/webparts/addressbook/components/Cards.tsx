import "./Addressbook.scss";
import * as React from "react";
import { Component } from "react";
import { Contact } from "../api provider/api_provider";

class ContactCard extends Component<{
  contact: Contact;
  showContact: Function;
}> {
  constructor(props: any) {
    super(props);
    this.showContact = this.showContact.bind(this);
  }

  showContact(e: React.MouseEvent) {
    this.props.showContact(e.currentTarget.id);
  }

  render() {
    return (
      <div
        className="contactCard"
        onClick={this.showContact}
        id={this.props.contact.key}
      >
        <p className="contactName">{this.props.contact.name}</p>
        <p className="contactDetails">{this.props.contact.email}</p>
        <p className="contactDetails">{this.props.contact.mobile}</p>
      </div>
    );
  }
}

export default ContactCard;
