import {SPHttpClient} from '@microsoft/sp-http';

export interface ITestImportantLinksProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  
  numGroups : number;
  useList : boolean;
  siteUrl: string;
  spHttpClient: SPHttpClient;

  iconPicker1: string;
  iconPicker2: string;
  iconPicker3: string;
  iconPicker4: string;
  iconPicker5: string;
  iconPicker6: string;
  iconPicker7: string;
  iconPicker8: string;
  iconPicker9: string;
  iconPicker10: string;

  iconColour1: string;
  iconColour2: string;
  iconColour3: string;
  iconColour4: string;
  iconColour5: string;
  iconColour6: string;
  iconColour7: string;
  iconColour8: string;
  iconColour9: string;
  iconColour10: string;

  groupTitle1 : string;
  groupTitle2 : string;
  groupTitle3 : string;
  groupTitle4 : string;
  groupTitle5 : string;
  groupTitle6 : string;
  groupTitle7 : string;
  groupTitle8 : string;
  groupTitle9 : string;
  groupTitle10 : string;

  linksData1 : any[];
  linksData2 : any[];
  linksData3 : any[];
  linksData4 : any[];
  linksData5 : any[];
  linksData6 : any[];
  linksData7 : any[];
  linksData8 : any[];
  linksData9 : any[];
  linksData10 : any[];
}