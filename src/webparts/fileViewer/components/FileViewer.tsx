import * as React from 'react';
import styles from './FileViewer.module.scss';
import { IFileViewerState } from './IFileViewerState';
import { IFileViewerProps } from './IFileViewerProps';
import { IListItems } from './IListems';
import { BaseButton, PanelType, Text } from 'office-ui-fabric-react';
import { IFramePanel } from "@pnp/spfx-controls-react/lib/IFramePanel";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { SPFI, spfi } from '@pnp/sp';
import { getSP } from './pnpjsConfig';




export default class FileViewer extends React.Component<IFileViewerProps, IFileViewerState, {}> {
  private _sp: SPFI;

  public constructor(props) {
    super(props);
    this.state = { ListItems: [], show: false, docUrl: "" };
    this._sp = getSP();
  }


  private async GetItems() {
    try {

      const spCache = spfi(this._sp);
      const response: IListItems[] = await spCache.web.lists
        .getByTitle('Document List')
        .items
        .select("Url", "Title", "Category", "SortOrder").orderBy("Category",true).orderBy("SortOrder",true)();

      console.log("response");
      console.log(response);


      this.setState({ ListItems: response });


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

