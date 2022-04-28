import * as React from 'react';
import styles from './FileViewer.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { IFileViewerState } from './IFileViewerState';
import { IFileViewerProps } from './IFileViewerProps';
import { SPHttpClient } from "@microsoft/sp-http";
import { IListItems } from './IListems';

export default class FileViewer extends React.Component<IFileViewerProps, IFileViewerState, {}> {


  public constructor(props) {
    super(props);
    this.state = { ListItems: [] };
  }


  private async GetItems() {
    try {
      var redirectionEmailURL =
        this.props.siteUrl +
        "/_api/web/lists/getbytitle('Document List')/Items?$select=Title" +
        ",Url" +
        ",Category" +
        ",SortOrder" ;

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
        
        <div>
          {this.state.ListItems.map((item)=>{
return (<div>{item.Title}</div>)
          })}
        </div>
        <div>

        </div>
      </section>
    );
  }
}
