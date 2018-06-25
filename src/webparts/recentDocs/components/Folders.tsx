import * as React from 'react';
import styles from './RecentDocs.module.scss';
import {
  IRecentDocsProps,
  IRecentDocsState,
  IDocListProps,
  IDocItemProps,
} from '../Model/IRecentDocsProps';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import DocList from './DocList';
import DocItem from './DocItem';

export default class RecentDocs extends React.Component<IRecentDocsProps, IRecentDocsState> {

  constructor(props: IRecentDocsProps, state: IRecentDocsState) {
    super(props);

    this.state = {
      items: [],
      parent: "",
      isLoaded: false,
    }
  }

  private testFolders =
    {
      root: {
        parent: "",
        folders: [
          { title: "Folder01", url: "", created: "2018-1-10", createdBy: "Bill Wang" },
          { title: "Folder02", url: "", created: "2018-1-10", createdBy: "Bill Wang" },
        ]
      }, Folder01: {
        parent: "Folder01",
        folders: [
          { title: "Folder03", url: "https://www.bing.com", created: "2018-1-10", createdBy: "Bill Wang" },
          { title: "Folder04", url: "https://www.bing.com", created: "2018-1-10", createdBy: "Bill Wang" },
        ]
      }, Folder02: {
        parent: "Folder02",
        folders: [
          { title: "Folder05", url: "https://www.bing.com", created: "2018-1-10", createdBy: "Bill Wang" },
          { title: "Folder06", url: "https://www.bing.com", created: "2018-1-10", createdBy: "Bill Wang" },
        ]
      }
    }
  private icon = "https://spoprod-a.akamaihd.net/files/odsp-next-prod_2018-06-08-sts_20180613.001/odsp-media/images/itemtypes/20/folder.svg";

  public render(): React.ReactElement<IRecentDocsProps> {
    // let serverUrl = this.props.context.pageContext.site.absoluteUrl.substring(0, this.props.context.pageContext.site.absoluteUrl.length-this.props.context.pageContext.site.serverRelativeUrl.length)
    // let moreUrl = `${serverUrl}/search/Pages/results.aspx?u=${this.props.context.pageContext.site.absoluteUrl}#Default={"k":"(IsDocument=\"True\" OR contentclass:\"STS_ListItem\")(FileExtension:doc OR FileExtension:docx OR FileExtension:xls OR FileExtension:xlsx OR FileExtension:ppt OR FileExtension:pptx OR FileExtension:pdf)", "o":[{"d":1,"p":"LastModifiedTime"}]}`

    const loadingElement: JSX.Element = <div style={{ 'margin-top': '40%' }}><Spinner label={'Loading Recent Documents...'} /></div>;
    let folders = this.LoadTestData();
    const showBack = this.state.parent == ""?"none":"block";

    return (
      <div className={styles.recentDocs}>
        <div className={styles.header}>
          <img className={styles.headerimg} src={String(require('../../../assets/images/title_recent_documents.png'))}></img>
          <span className={styles.headertitle}>{this.props.listTitle}</span>
        </div>
        <div className={styles.container} >
          <div className={styles.back} style={{ 'display': showBack }}>
            <a onClick={this.Back.bind(this)}>{"<- Back"}</a>
          </div>
          <ul className={styles.list} >
            {folders}
          </ul>
        </div>
      </div>
    );
  }

  private Back() {
    this.setState({
      parent: "",
    });
  }

  private Browse(arg) {
    if (this.testFolders[this.state.parent] != undefined) {
      document.location.href = arg;
      return;
    }
    this.setState({
      parent: arg,
    });
  }

  private LoadTestData() {
    let list = [];
    let count = 1;
    if (this.state.parent == "") {
      this.testFolders.root.folders.forEach(element => {
        list.push(
          this.BuildItem(element, count)
        );
        count++;
      });
    }
    else {
      this.testFolders[this.state.parent].folders.forEach(element => {
        list.push(
          this.BuildItem(element, count)
        );
        count++;
      });
    }
    return list;
  }



  private BuildItem(element, count) {
    let folderLink;

    if (element.url == "") {
      folderLink = <a onClick={this.Browse.bind(this, element.title)}>{element.title}</a>
    }
    else {
      folderLink = <a onClick={this.Browse.bind(this, element.url)}>{element.title}</a>
    }

    return <li className={styles.item} key={count}>
      <table>
        <tbody>
          <tr>
            <td className={styles.icon}>
              <img src={this.icon}></img>
            </td>
            <td className={styles.doc}>
              <table>
                <tbody>
                  <tr>
                    <td>
                      {folderLink}
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <span>Modified Date: {element.created}</span>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <span>By: {element.createdBy}</span>
                    </td>
                  </tr>
                </tbody>
              </table>
            </td>
          </tr>
        </tbody>
      </table>
    </li>;
  }
}
