import * as React from 'react';
import styles from './FileViewer.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { IFileViewerState } from './IFileViewerState';
import { IFileViewerProps } from './IFileViewerProps';
import { SPHttpClient } from "@microsoft/sp-http";
import { IListItems } from './IListems';
import { BaseButton, PanelType, Text } from 'office-ui-fabric-react';
import { IFramePanel } from "@pnp/spfx-controls-react/lib/IFramePanel";

export default class FileViewer extends React.Component<IFileViewerProps, IFileViewerState, {}> {


  public constructor(props) {
    super(props);
    this.state = { ListItems: [], show: false, docUrl: "" };
  }


  private async GetItems() {
    try {
      var redirectionEmailURL =
        this.props.siteUrl +
        "/_api/web/lists/getbytitle('Document List')/Items?$select=Title" +
        ",Url" +
        ",Category" +
        ",SortOrder";

      const responseEmail = await this.props.context.spHttpClient.get(
        redirectionEmailURL,
        SPHttpClient.configurations.v1
      );
      var returnValues: IListItems[] = [];
      const responseEmailJSON = await responseEmail.json();
      if (responseEmailJSON.value !== null) {
        var resultJSONArray = responseEmailJSON.value;

        //  resultJSONArray = [
        //  {Title:"Ashish",SortOrder:1,Category:"Developer",Url:"http://google.com"},
        //  {Title:"Ghatak",SortOrder:2,Category:"Developer",Url:"http://google.com"},
        //]

        resultJSONArray.map((att) => {
          returnValues.push({
            Title: att.Title,
            Category: att.Category,
            Url: att.Url,
            LinkTitle: att.LinkTitle
          });
        });
        this.setState({ ListItems: returnValues });
        //  ListItems = [
        //  {Title:"Ashish",SortOrder:1,Category:"Developer",Url:"http://google.com"},
        //  {Title:"Ghatak",SortOrder:2,Category:"Developer",Url:"http://google.com"},
        //]
      }

    } catch (error) {
      console.log("Error in GetItem : " + error);
    }
  }

  public componentDidMount(): void {
    this.GetItems();

  }

  public _OnDismiss() {
    this.setState({ show: false });
  }


  public _OnItemClick(url, ev) {
    this.setState({ show: true, docUrl: url });
  }

  public render(): React.ReactElement<IFileViewerProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.fileViewer} ${hasTeamsContext ? styles.teams : ''}`}>

        <div className={styles.Navigation}>
          <h1>Navigation</h1>
          {this.state.ListItems.map((item) => {
            return (<h5 onClick={this._OnItemClick.bind(this, item.Url)}>{item.Title}</h5>)
          })}
        </div>
        <div className={styles.FileLoadViewer}>
          <h1>File Viewer</h1>
          <IFramePanel url={this.state.docUrl}
            type={PanelType.extraLarge}
            headerText="Panel Title"
            closeButtonAriaLabel="Close"
            isOpen={this.state.show}
            onDismiss={this._OnDismiss.bind(this)}
          />

        </div>
      </section>
    );
  }
}

