import "./Addressbook.scss";
import * as React from 'react';
import { Component } from "react";
import blogIcon from "../assets/blog-icon.png";

class NavBar extends Component<{ handleDialog: Function }> {
  constructor(props: any) {
    super(props);
    this.handleAdd = this.handleAdd.bind(this);
  }

  handleAdd(e: React.MouseEvent) {
    e.preventDefault();
    this.props.handleDialog();
  }

  render() {
    return (
      <ul className="bar">
        <li>
          <a href="index.html">HOME</a>
        </li>
        <li>
          <a onClick={this.handleAdd} href="">
            +ADD
          </a>
        </li>

        <li className="icon">
          <img className="icon" src={blogIcon} alt="Blog Icon" />
        </li>
      </ul>
    );
  }
}

export default NavBar;
