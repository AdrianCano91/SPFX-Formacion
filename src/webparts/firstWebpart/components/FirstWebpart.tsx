import * as React from 'react';
import styles from './FirstWebpart.module.scss';
import { IFirstWebpartProps } from './IFirstWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ISPLists } from '../Interfaces/ISPLists';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import { ISPList } from '../Interfaces/ISPList';
import { getSP } from "../pnpjsConfig";
import { SPFI } from '@pnp/sp';
import "@pnp/sp/sites";


export default class FirstWebpart extends React.Component<IFirstWebpartProps, {items: any[]}> {
  private _sp: SPFI;

  constructor(props: IFirstWebpartProps) {
    super(props);
    this.state = {
        items: [],       
      };
      this._sp = getSP();
    debugger;
}

  public componentDidMount(): void {

    this._getListData();
  
  }

  private async _getListData() {
    const lists = await this._sp.web.lists();
    console.log("lists", lists)
    this.setState({ items: lists});
  }

  private _updateState(value: any[]) {
    this.setState({ items: value});
  }

  public render(): JSX.Element {
    var listado = [];
    listado = this.state.items;
    return (
      <section className={`${styles.firstWebpart} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={this.props.isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(this.props.userDisplayName)}!</h2>
          <div>{this.props.environmentMessage}</div>
          <div>Web part property value: <strong>{escape(this.props.description)}</strong></div>
          <p>{escape(this.props.test)}</p>
          <p>{this.props.test1}</p>
          <p>{escape(this.props.test2)}</p>
          <p>{this.props.test3}</p>
          {listado.map((element) => {
            
            
            return <ul className={styles.list}>
              <li className={styles.listItem}>
                <span className="ms-font-l">{element.Title}</span>
              </li>
            </ul>
            
          })}
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <div>Web part description: <strong>{escape(this.props.description)}</strong></div>
          <div>Web part test: <strong>{escape(this.props.test)}</strong></div>
          <div>Loading from: <strong>{escape(this.props.context.pageContext.web.title)}</strong></div>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
      </section>
    );
  }
}
