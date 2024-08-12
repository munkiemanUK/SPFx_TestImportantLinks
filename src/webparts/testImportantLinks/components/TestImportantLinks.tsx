import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './TestImportantLinks.module.scss';
import type { ITestImportantLinksProps } from './ITestImportantLinksProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { Icon } from '@fluentui/react';
//import { SPFI } from '@pnp/sp';
import { Web } from "@pnp/sp/webs"; 
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { getSP } from '../pnpjsConfig';
import {IColumn} from '@fluentui/react';
import { SPHttpClient} from '@microsoft/sp-http';

let panelHTML: any[] = [];
let links: any[] = [];

export interface IListItem {
  linkTitle: string;
  linkURL: string;
  linkOrder: number;
  linkBrowse: string;
  linkGroupID: number;
  linkGroupName: string;
}

const TestImportantLinks: React.FC<ITestImportantLinksProps> = (props) => {

  const {
  //  description,
  //  isDarkTheme,
  //  environmentMessage,
  //  hasTeamsContext,
  //  userDisplayName
    siteUrl,
    spHttpClient
  } = props;

  const [webPartTitle, setWebPartTitle] = useState('');
  const [grouptitle1, setGroupTitle1] = useState('');
  const [grouptitle2, setGroupTitle2] = useState('');
  const [grouptitle3, setGroupTitle3] = useState('');
  const [grouptitle4, setGroupTitle4] = useState('');
  const [grouptitle5, setGroupTitle5] = useState('');
  const [grouptitle6, setGroupTitle6] = useState('');
  const [grouptitle7, setGroupTitle7] = useState('');
  const [grouptitle8, setGroupTitle8] = useState('');
  const [grouptitle9, setGroupTitle9] = useState('');
  const [grouptitle10, setGroupTitle10] = useState('');

  const [listItems, setListItems] = useState<IListItem[]>([]);
  //const [listFlag, setListFlag] = useState<boolean>(false);
  //const [columns, setColumns] = useState<IColumn[]>([
  const columns: IColumn[] = [
    {
      key: "linkTitle",
      name: "",
      fieldName: "LinkName",
      minWidth: 0,
      maxWidth: 50,
      isResizable: true,
      data: "string",
      isPadded: true
    },
    {
      key: "linkURL",
      name: "",
      fieldName: "LinkURL",
      minWidth: 0,
      maxWidth: 50,
      isResizable: true,
      data: "string",
      isPadded: true
    },
    {
      key: "linkBrowse",
      name: "",
      fieldName: "LinkBrowse",
      minWidth: 0,
      maxWidth: 50,
      isResizable: true,
      data: "string",
      isPadded: true
    },
    {
      key: "linkOrder",
      name: "",
      fieldName: "LinkOrder",
      minWidth: 0,
      maxWidth: 50,
      isResizable: true,
      data: "number",
      isPadded: true
    },
    {
      key: "linkGroupID",
      name: "",
      fieldName: "GroupID",
      minWidth: 0,
      maxWidth: 50,
      isResizable: true,
      data: "number",
      isPadded: true
    },
    {
      key: "linkGroupName",
      name: "",
      fieldName: "GroupName",
      minWidth: 0,
      maxWidth: 50,
      isResizable: true,
      data: "number",
      isPadded: true
    }
  ]
  //]);

  const _sp = getSP();

  useEffect(() => {
    if (props.useList) {
      _getListData();
    }
  }, [props.useList]);

  useEffect(() => {
    if (props.useList) {
      _renderListData();
    }
  }, [listItems]);

  useEffect(() => {
    console.log("componentDidMount called.");
    console.log("Fetching CanvasContent1 WebPartData Title.");
    (async () => {
      const endpoint = `${siteUrl}/_api/sitepages/pages(1)?$select=CanvasContent1&expand=CanvasContent1`;
      const rawResponse = await spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      const jsonResponse = await rawResponse.json();
      const jsonCanvasContent = jsonResponse.CanvasContent1;
      const parseCanvasContent = JSON.parse(jsonCanvasContent);
      //const webpartData = canvasContent.webPartData;
      //console.log("canvascontent json",jsonCanvasContent);
      console.log("canvascontent parse",parseCanvasContent);
      //parseCanvasContent.forEach((item:any,index:number) => {
      //let index : number = 0;
      for(const item of parseCanvasContent){
        console.log("webPartData Title",item.webPartData.title);
        let itemTitle : string = item.webPartData.title;
        let itemGroup1 : string = item.webPartData.properties.Group1Title;
        let itemGroup2 : string = item.webPartData.properties.Group2Title;
        let itemGroup3 : string = item.webPartData.properties.Group3Title;
        let itemGroup4 : string = item.webPartData.properties.Group4Title;
        let itemGroup5 : string = item.webPartData.properties.Group5Title;
        let itemGroup6 : string = item.webPartData.properties.Group6Title;
        let itemGroup7 : string = item.webPartData.properties.Group7Title;
        let itemGroup8 : string = item.webPartData.properties.Group8Title;
        let itemGroup9 : string = item.webPartData.properties.Group9Title;
        let itemGroup10 : string = item.webPartData.properties.Group10Title;

        if(item.webPartData.title === "Important Links"){
          setWebPartTitle(itemTitle);
          setGroupTitle1(itemGroup1);
          setGroupTitle2(itemGroup2);
          setGroupTitle3(itemGroup3);
          setGroupTitle4(itemGroup4);
          setGroupTitle5(itemGroup5);
          setGroupTitle6(itemGroup6);
          setGroupTitle7(itemGroup7);
          setGroupTitle8(itemGroup8);
          setGroupTitle9(itemGroup9);
          setGroupTitle10(itemGroup10);
          //break;          
        }
        //index++;
      }
      console.log(webPartTitle,grouptitle1,grouptitle2,grouptitle3,grouptitle4,grouptitle5,grouptitle6,grouptitle7,grouptitle8,grouptitle9,grouptitle10);

      //})
    })();
  },[siteUrl,spHttpClient]);

  const _getListData = async () => {
    const data: IListItem[] = [];
    const view = `<View>
                    <Query>
                      <OrderBy>
                        <FieldRef Name="GroupID" Ascending="TRUE" />
                        <FieldRef Name="LinkOrder" Ascending="TRUE" />
                      </OrderBy>          
                    </Query>
                  </View>`;
    const web = Web([_sp.web, props.siteUrl]);
    const response = await web.lists.getByTitle('Important Links').getItemsByCAMLQuery({ ViewXml: view });
    console.log("camlItems", response);
    response.forEach((item: { LinkName: any; LinkURL: any; LinkOrder: any; LinkBrowse: any; GroupID: any; Title: any }) => {
      console.log(item.LinkName);
      data.push({
        linkTitle: item.LinkName,
        linkURL: item.LinkURL,
        linkOrder: item.LinkOrder,
        linkBrowse: item.LinkBrowse,
        linkGroupID: item.GroupID,
        linkGroupName: item.Title
      });
    });
    console.log("data", data,columns);
    setListItems(data);
  };

  const _renderListData = () => {
    let linkHTML: string = '';

    listItems.forEach(item => {
      let linkGroupId: number = Math.floor(item.linkGroupID);
      const link: Element = document.querySelector('#linkContainer' + linkGroupId)!;

      console.log("GroupID=" + linkGroupId + " Name=" + item.linkTitle);
      linkHTML = `<div class="row linkrow"><a href="${item.linkURL}" target="${item.linkBrowse}">
                    <h5 class="">${item.linkTitle}</h5>
                  </a></div>`;
      if (link) { link.innerHTML += linkHTML };
    });
  };

  const AddPanel = () => {
    let numAccordions = Number(props.numGroups);
    let groupTitle: string = "";
    let groupContainerId: string = "";
    let groupHeadingId: string = "";
    let linksAccordionId: string = "";
    let linksAccordionHash: string = "";
    let linkContainerId: string = "";
    let iconName: any;
    let iconColour: string = "";
    panelHTML = [];

    if (props.numGroups !== undefined) {
      for (let i = 1; i <= numAccordions; i++) {
        groupContainerId = "groupContainer" + i;
        groupHeadingId = "groupHeader" + i;
        linksAccordionId = "linksAccordion" + i;
        linksAccordionHash = "#linksAccordion" + i;
        linkContainerId = "linkContainer" + i;

        switch (i) {
          case 1:
            groupTitle = props.groupTitle1;
            iconName = props.iconPicker1;
            iconColour = props.iconColour1;
            links = props.linksData1;
            break;
          case 2:
            groupTitle = props.groupTitle2;
            iconName = props.iconPicker2;
            iconColour = props.iconColour2;
            links = props.linksData2;
            break;
          case 3:
            groupTitle = props.groupTitle3;
            iconName = props.iconPicker3;
            iconColour = props.iconColour3;
            links = props.linksData3;
            break;
          case 4:
            groupTitle = props.groupTitle4;
            iconName = props.iconPicker4;
            iconColour = props.iconColour4;
            links = props.linksData4;
            break;
          case 5:
            groupTitle = props.groupTitle5;
            iconName = props.iconPicker5;
            iconColour = props.iconColour5;
            links = props.linksData5;
            break;
          case 6:
            groupTitle = props.groupTitle6;
            iconName = props.iconPicker6;
            iconColour = props.iconColour6;
            links = props.linksData6;
            break;
          case 7:
            groupTitle = props.groupTitle7;
            iconName = props.iconPicker7;
            iconColour = props.iconColour7;
            links = props.linksData7;
            break;
          case 8:
            groupTitle = props.groupTitle8;
            iconName = props.iconPicker8;
            iconColour = props.iconColour8;
            links = props.linksData8;
            break;
          case 9:
            groupTitle = props.groupTitle9;
            iconName = props.iconPicker9;
            iconColour = props.iconColour9;
            links = props.linksData9;
            break;
          case 10:
            groupTitle = props.groupTitle10;
            iconName = props.iconPicker10;
            iconColour = props.iconColour10;
            links = props.linksData10;
            break;
        }

        panelHTML.push(
          <div className="accordion-item" id={groupContainerId}>
            <h3 className="accordion-header" id={groupHeadingId}>
              <button className="accordion-button collapsed" role="button" data-bs-toggle="collapse" data-bs-target={linksAccordionHash} aria-expanded="false" aria-controls="collapse">
                <div className="col-1"><Icon style={{ fontSize: '24px', color: iconColour }} iconName={iconName} className="ms-IconExample me-2" /></div>
                <div className="col">{groupTitle}</div>
              </button>
            </h3>
            <div id={linksAccordionId} className="accordion-collapse collapse" data-bs-parent="#linkAccordion">
              <div className="accordion-body" id={linkContainerId}>
                <div className="row">
                  {links && links.map((val) => {
                    return (<div><span>{val.LinkTitle}</span><span style={{ marginLeft: 10 }}>{val.LinkURL}</span></div>);
                  })}
                </div>
              </div>
            </div>
          </div>
        );
      }
    }
  };

  AddPanel();

  return (
    <section className={`${styles.testImportantLinks} ${props.hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <h2 className="welcome">{props.description}</h2>
        <div className="accordion accordion-flush" id="linkAccordion">
          {panelHTML}
        </div>
      </div>
    </section>
  );
};

export default TestImportantLinks;