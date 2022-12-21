import "../Addressbook.scss";
import * as React from 'react';
import { Component } from "react";

class Header extends Component {
  render() {
    return (
      <div className="top">
        <p id="heading">Address Book</p>
      </div>
    );
  }
}

export default Header;
