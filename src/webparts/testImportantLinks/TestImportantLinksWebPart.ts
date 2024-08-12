import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneCheckbox,
  PropertyPaneButton,
  PropertyPaneLabel,
  PropertyPaneButtonType,  
  IPropertyPaneGroup,
  //PropertyPaneDropdown,
  //PropertyPaneHorizontalRule

} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TestImportantLinksWebPartStrings';
import TestImportantLinks from './components/TestImportantLinks';
import { ITestImportantLinksProps } from './components/ITestImportantLinksProps';
import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { getSP } from './pnpjsConfig';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';

require('bootstrap');
require('./styles/custom.css');

let linkList : any={};
let iconCode : any={};
let iconColour : any={};
let linksCode : any={};

export interface ITestImportantLinksWebPartProps {
  description: string;
  numGroups : number;
  useList : boolean;
  siteUrl: string;

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

  iconPicker1: any;
  iconPicker2: any;
  iconPicker3: any;
  iconPicker4: any;
  iconPicker5: any;
  iconPicker6: any;
  iconPicker7: any;
  iconPicker8: any;
  iconPicker9: any;
  iconPicker10: any;

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

  linksGroup1 : any[];
  linksGroup2 : any[];
  linksGroup3 : any[];
  linksGroup4 : any[];
  linksGroup5 : any[];
  linksGroup6 : any[];
  linksGroup7 : any[];
  linksGroup8 : any[];
  linksGroup9 : any[];
  linksGroup10 : any[];
}

export default class TestImportantLinksWebPart extends BaseClientSideWebPart<ITestImportantLinksWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ITestImportantLinksProps> = React.createElement(
      TestImportantLinks,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        numGroups : this.properties.numGroups,
        useList : this.properties.useList,  
        siteUrl: this.context.pageContext.site.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        //context: this.context,     

        iconPicker1: this.properties.iconPicker1,
        iconPicker2: this.properties.iconPicker2,
        iconPicker3: this.properties.iconPicker3,
        iconPicker4: this.properties.iconPicker4,
        iconPicker5: this.properties.iconPicker5,
        iconPicker6: this.properties.iconPicker6,
        iconPicker7: this.properties.iconPicker7,
        iconPicker8: this.properties.iconPicker8,
        iconPicker9: this.properties.iconPicker9,
        iconPicker10: this.properties.iconPicker10,
        iconColour1 : this.properties.iconColour1,
        iconColour2 : this.properties.iconColour2,
        iconColour3 : this.properties.iconColour3,
        iconColour4 : this.properties.iconColour4,
        iconColour5 : this.properties.iconColour5,
        iconColour6 : this.properties.iconColour6,
        iconColour7 : this.properties.iconColour7,
        iconColour8 : this.properties.iconColour8,
        iconColour9 : this.properties.iconColour9,
        iconColour10 : this.properties.iconColour10,
        groupTitle1: this.properties.groupTitle1,
        groupTitle2: this.properties.groupTitle2,
        groupTitle3: this.properties.groupTitle3,
        groupTitle4: this.properties.groupTitle4,
        groupTitle5: this.properties.groupTitle5,
        groupTitle6: this.properties.groupTitle6,
        groupTitle7: this.properties.groupTitle7,
        groupTitle8: this.properties.groupTitle8,
        groupTitle9: this.properties.groupTitle9,
        groupTitle10: this.properties.groupTitle10,
        linksData1: this.properties.linksGroup1,
        linksData2: this.properties.linksGroup2,
        linksData3: this.properties.linksGroup3,
        linksData4: this.properties.linksGroup4,
        linksData5: this.properties.linksGroup5,
        linksData6: this.properties.linksGroup6,
        linksData7: this.properties.linksGroup7,
        linksData8: this.properties.linksGroup8,
        linksData9: this.properties.linksGroup9,
        linksData10: this.properties.linksGroup10
      }
    );

    ReactDom.render(element, this.domElement);
    if(this.properties.useList){
      this._renderDataAsync();
    }
  }

  private _renderDataAsync() : void {
    this._getData()
    .then((response) => {
      this._renderData(response);
    });
  }

  private async _getData() : Promise<any> {
    const Uri = this.context.pageContext.site.absoluteUrl + `/_api/sitepages/pages(1)?$select=CanvasContent1&expand=CanvasContent1`; //`/_api/web/lists/getbytitle('Site%20Pages')/items(1)/FieldValuesAsHTML`;
    console.log("Uri",Uri);
    return await this.context.spHttpClient.get(Uri, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
  } 

  private _renderData(items:any): void {
    //let id = this.context.pageContext.listItem?.id;
    const canvasContent = JSON.parse(items.CanvasContent1)

    //console.log("items",items);
    console.log("group1",canvasContent[8].id);
    //console.log("canvascontent",canvasContent);

    canvasContent.forEach((item:any,index:number)=>{
      if(item.webPartData.title !== undefined){
        let wpTitle : string = item.webPartData.title;
        if(wpTitle === "Important Links"){        
          this.properties.groupTitle1 = item.webPartData.properties.Group1Title;
          this.properties.iconPicker1 = "Link12";
          this.properties.iconColour1 = "black";
          this.properties.groupTitle2 = item.webPartData.properties.Group2Title;
          this.properties.iconPicker2 = "Link12";
          this.properties.iconColour2 = "black";
          this.properties.groupTitle3 = item.webPartData.properties.Group3Title;
          this.properties.iconPicker3 = "Link12";
          this.properties.iconColour3 = "black";
          this.properties.groupTitle4 = item.webPartData.properties.Group4Title;
          this.properties.iconPicker4 = "Link12";
          this.properties.iconColour4 = "black";
          this.properties.groupTitle5 = item.webPartData.properties.Group5Title;
          this.properties.iconPicker5 = "Link12";
          this.properties.iconColour5 = "black";
          this.properties.groupTitle6 = item.webPartData.properties.Group6Title;
          this.properties.iconPicker6 = "Link12";
          this.properties.iconColour6 = "black";
          this.properties.groupTitle7 = item.webPartData.properties.Group7Title;
          this.properties.iconPicker7 = "Link12";
          this.properties.iconColour7 = "black";
          this.properties.groupTitle8 = item.webPartData.properties.Group8Title;
          this.properties.iconPicker8 = "Link12";
          this.properties.iconColour8 = "black";
          this.properties.groupTitle9 = item.webPartData.properties.Group9Title;
          this.properties.iconPicker9 = "Link12";
          this.properties.iconColour9 = "black";
          this.properties.groupTitle10 = item.webPartData.properties.Group10Title;
          this.properties.iconPicker10 = "Link12";
          this.properties.iconColour10 = "black";
          this.properties.numGroups = item.webPartData.properties.Slider;
        }

        console.log("canvasContent Item",item.webPartData.title);
        console.log("canvascontent",canvasContent[index]);
        console.log("instanceID",this.context.instanceId);
      }
    })
  }

  public async onInit(): Promise<void> {
    await super.onInit();
    getSP(this.context);

    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.3/font/bootstrap-icons.css");

    return this._getEnvironmentMessage().then(message => {
      //this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private buttonClick(): void {  
    const currentWebUrl = this.context.pageContext.web.absoluteUrl; 
    window.open(currentWebUrl+'/Lists/ImportantLinks/AllItems.aspx','_blank');  
    //return "test"  
  }

  protected textBoxValidationMethod(value: string): string {
    if (value.length < 10 || value.length > 50) 
    {
      return "App name should be at least 10 but less than 50 characters!"; 
    }
    else 
    { 
      return ""; 
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const groupPanels: IPropertyPaneGroup[] =[];
    const linkPanels: IPropertyPaneGroup[] =[];

    //if(this.properties.useList){alert('use list')}

    if(this.properties.numGroups>0){
      for(let x=1; x<=this.properties.numGroups;x++){

        switch(x){
          case 1:
            if(this.properties.groupTitle1 === undefined){this.properties.groupTitle1 = `Link Group ${x}`;}

            iconCode = PropertyFieldIconPicker('iconPicker1', {
                currentIcon: this.properties.iconPicker1,
                key: "iconPickerId",
                onSave: (icon: string) => { console.log(icon); this.properties.iconPicker1 = icon;},
                onChanged:(icon: string) => { console.log(icon);  },
                buttonLabel: "Select Icon",
                renderOption: "panel",
                properties: this.properties,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                label: "Icon Picker for Group "+x            
              });
            iconColour = PropertyFieldColorPicker('iconColour1', {
                label: 'Icon Colour',
                selectedColor: this.properties.iconColour1,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                properties: this.properties,
                disabled: false,
                debounce: 1000,
                isHidden: false,
                alphaSliderHidden: false,
                style: PropertyFieldColorPickerStyle.Full,
                iconName: 'Precipitation',
                key: 'colorFieldId'
              });

            if(!this.properties.useList){ 
              linksCode = PropertyFieldCollectionData("linksGroup1", {
                key: "linksGroup1",
                label: "Group 1 Links",
                panelHeader: "Enter the links for Group "+x,
                manageBtnLabel: "Manage Links Group "+x,
                value: this.properties.linksGroup1,
                fields: [
                  {
                    id: "LinkTitle",
                    title: "Link Title",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkURL",
                    title: "Link URL",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkBrowse",
                    title: "Link Browse",
                    type: CustomCollectionFieldType.dropdown,
                    options: [
                      {
                        key: "_self",
                        text: "Current Tab"
                      },
                      {
                        key: "_blank",
                        text: "New Tab"
                      },
                      {
                        key: "_parent",
                        text: "Current Browser - New Window"
                      },
                      {
                        key: "_top",
                        text: "New Browser"
                      }                      
                    ],
                    required: true
                  }
                ],
                disabled: false
              })
            }                                      
            break;
          case 2:
            if(this.properties.groupTitle2 === undefined){this.properties.groupTitle2 = `Link Group ${x}`;}

            iconCode = PropertyFieldIconPicker('iconPicker2', {
                        currentIcon: this.properties.iconPicker2,
                        key: "iconPickerId",
                        onSave: (icon: string) => { console.log(icon); this.properties.iconPicker2 = icon;},
                        onChanged:(icon: string) => { console.log(icon);  },
                        buttonLabel: "Select Icon",
                        renderOption: "panel",
                        properties: this.properties,
                        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                        label: "Icon Picker for Group "+x            
                      });

            iconColour = PropertyFieldColorPicker('iconColour2', {
                          label: 'Icon Colour',
                          selectedColor: this.properties.iconColour2,
                          onPropertyChange: this.onPropertyPaneFieldChanged,
                          properties: this.properties,
                          disabled: false,
                          debounce: 1000,
                          isHidden: false,
                          alphaSliderHidden: false,
                          style: PropertyFieldColorPickerStyle.Full,
                          iconName: 'Precipitation',
                          key: 'colorFieldId'
                        });

            if(!this.properties.useList){ 
              linksCode = PropertyFieldCollectionData("linksGroup2", {
                key: "linksGroup2",
                label: "Group 2 Links",
                panelHeader: "Enter the links for Group "+x,
                manageBtnLabel: "Manage Links Group "+x,
                value: this.properties.linksGroup2,
                fields: [
                  {
                    id: "LinkTitle",
                    title: "Link Title",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkURL",
                    title: "Link URL",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkBrowse",
                    title: "Link Browse",
                    type: CustomCollectionFieldType.dropdown,
                    options: [
                      {
                        key: "_self",
                        text: "Current Tab"
                      },
                      {
                        key: "_blank",
                        text: "New Tab"
                      },
                      {
                        key: "_parent",
                        text: "Current Browser - New Window"
                      },
                      {
                        key: "_top",
                        text: "New Browser"
                      }                      
                    ],
                    required: true
                  }
                ],
                disabled: false
              })
            }
            break;
          case 3:
            if(this.properties.groupTitle3 === undefined){this.properties.groupTitle3 = `Link Group ${x}`;}
            iconCode = PropertyFieldIconPicker('iconPicker3', {
                        currentIcon: this.properties.iconPicker3,
                        key: "iconPickerId",
                        onSave: (icon: string) => { console.log(icon); this.properties.iconPicker3 = icon;},
                        onChanged:(icon: string) => { console.log(icon);  },
                        buttonLabel: "Select Icon",
                        renderOption: "panel",
                        properties: this.properties,
                        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                        label: "Icon Picker for Group "+x            
                      });

            iconColour = PropertyFieldColorPicker('iconColour3', {
                        label: 'Icon Colour',
                        selectedColor: this.properties.iconColour3,
                        onPropertyChange: this.onPropertyPaneFieldChanged,
                        properties: this.properties,
                        disabled: false,
                        debounce: 1000,
                        isHidden: false,
                        alphaSliderHidden: false,
                        style: PropertyFieldColorPickerStyle.Full,
                        iconName: 'Precipitation',
                        key: 'colorFieldId'
                      });
        
            if(!this.properties.useList){ 
              linksCode = PropertyFieldCollectionData("linksGroup3", {
                key: "linksGroup3",
                label: "Group 3 Links",
                panelHeader: "Enter the links for Group "+x,
                manageBtnLabel: "Manage Links Group "+x,
                value: this.properties.linksGroup3,
                fields: [
                  {
                    id: "LinkTitle",
                    title: "Link Title",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkURL",
                    title: "Link URL",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkBrowse",
                    title: "Link Browse",
                    type: CustomCollectionFieldType.dropdown,
                    options: [
                      {
                        key: "_self",
                        text: "Current Tab"
                      },
                      {
                        key: "_blank",
                        text: "New Tab"
                      },
                      {
                        key: "_parent",
                        text: "Current Browser - New Window"
                      },
                      {
                        key: "_top",
                        text: "New Browser"
                      }                      
                    ],
                    required: true
                  }
                ],
                disabled: false
              })
            }                                      

            break;
          case 4:
            if(this.properties.groupTitle4 === undefined){this.properties.groupTitle4 = `Link Group ${x}`;}
            iconCode = PropertyFieldIconPicker('iconPicker4', {
              currentIcon: this.properties.iconPicker4,
              key: "iconPickerId",
              onSave: (icon: string) => { console.log(icon); this.properties.iconPicker4 = icon;},
              onChanged:(icon: string) => { console.log(icon);  },
              buttonLabel: "Select Icon",
              renderOption: "panel",
              properties: this.properties,
              onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
              label: "Icon Picker for Group "+x            
            });
            iconColour = PropertyFieldColorPicker('iconColour4', {
                        label: 'Icon Colour',
                        selectedColor: this.properties.iconColour4,
                        onPropertyChange: this.onPropertyPaneFieldChanged,
                        properties: this.properties,
                        disabled: false,
                        debounce: 1000,
                        isHidden: false,
                        alphaSliderHidden: false,
                        style: PropertyFieldColorPickerStyle.Full,
                        iconName: 'Precipitation',
                        key: 'colorFieldId'
                      });

            if(!this.properties.useList){ 
              linksCode = PropertyFieldCollectionData("linksGroup4", {
                key: "linksGroup4",
                label: "Group 4 Links",
                panelHeader: "Enter the links for Group "+x,
                manageBtnLabel: "Manage Links Group "+x,
                value: this.properties.linksGroup4,
                fields: [
                  {
                    id: "LinkTitle",
                    title: "Link Title",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkURL",
                    title: "Link URL",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkBrowse",
                    title: "Link Browse",
                    type: CustomCollectionFieldType.dropdown,
                    options: [
                      {
                        key: "_self",
                        text: "Current Tab"
                      },
                      {
                        key: "_blank",
                        text: "New Tab"
                      },
                      {
                        key: "_parent",
                        text: "Current Browser - New Window"
                      },
                      {
                        key: "_top",
                        text: "New Browser"
                      }                      
                    ],
                    required: true
                  }
                ],
                disabled: false
              })
            }                                      
          break;
          case 5:
            if(this.properties.groupTitle5 === undefined){this.properties.groupTitle5 = `Link Group ${x}`;}
            iconCode = PropertyFieldIconPicker('iconPicker5', {
              currentIcon: this.properties.iconPicker5,
              key: "iconPickerId",
              onSave: (icon: string) => { console.log(icon); this.properties.iconPicker5 = icon;},
              onChanged:(icon: string) => { console.log(icon);  },
              buttonLabel: "Select Icon",
              renderOption: "panel",
              properties: this.properties,
              onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
              label: "Icon Picker for Group "+x            
            });

            iconColour = PropertyFieldColorPicker('iconColour5', {
                        label: 'Icon Colour',
                        selectedColor: this.properties.iconColour5,
                        onPropertyChange: this.onPropertyPaneFieldChanged,
                        properties: this.properties,
                        disabled: false,
                        debounce: 1000,
                        isHidden: false,
                        alphaSliderHidden: false,
                        style: PropertyFieldColorPickerStyle.Full,
                        iconName: 'Precipitation',
                        key: 'colorFieldId'
                      });

            if(!this.properties.useList){ 
              linksCode = PropertyFieldCollectionData("linksGroup5", {
                key: "linksGroup5",
                label: "Group 5 Links",
                panelHeader: "Enter the links for Group "+x,
                manageBtnLabel: "Manage Links Group "+x,
                value: this.properties.linksGroup5,
                fields: [
                  {
                    id: "LinkTitle",
                    title: "Link Title",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkURL",
                    title: "Link URL",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkBrowse",
                    title: "Link Browse",
                    type: CustomCollectionFieldType.dropdown,
                    options: [
                      {
                        key: "_self",
                        text: "Current Tab"
                      },
                      {
                        key: "_blank",
                        text: "New Tab"
                      },
                      {
                        key: "_parent",
                        text: "Current Browser - New Window"
                      },
                      {
                        key: "_top",
                        text: "New Browser"
                      }                      
                    ],
                    required: true
                  }
                ],
                disabled: false
              })
            }                                      
            break;
          case 6:
            if(this.properties.groupTitle6 === undefined){this.properties.groupTitle6 = `Link Group ${x}`;}
            iconCode = PropertyFieldIconPicker('iconPicker6', {
              currentIcon: this.properties.iconPicker6,
              key: "iconPickerId",
              onSave: (icon: string) => { console.log(icon); this.properties.iconPicker6 = icon;},
              onChanged:(icon: string) => { console.log(icon);  },
              buttonLabel: "Select Icon",
              renderOption: "panel",
              properties: this.properties,
              onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
              label: "Icon Picker for Group "+x            
            });

            iconColour = PropertyFieldColorPicker('iconColour6', {
                        label: 'Icon Colour',
                        selectedColor: this.properties.iconColour6,
                        onPropertyChange: this.onPropertyPaneFieldChanged,
                        properties: this.properties,
                        disabled: false,
                        debounce: 1000,
                        isHidden: false,
                        alphaSliderHidden: false,
                        style: PropertyFieldColorPickerStyle.Full,
                        iconName: 'Precipitation',
                        key: 'colorFieldId'
                      });

            if(!this.properties.useList){ 
              linksCode = PropertyFieldCollectionData("linksGroup6", {
                key: "linksGroup6",
                label: "Group 6 Links",
                panelHeader: "Enter the links for Group "+x,
                manageBtnLabel: "Manage Links Group "+x,
                value: this.properties.linksGroup6,
                fields: [
                  {
                    id: "LinkTitle",
                    title: "Link Title",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkURL",
                    title: "Link URL",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkBrowse",
                    title: "Link Browse",
                    type: CustomCollectionFieldType.dropdown,
                    options: [
                      {
                        key: "_self",
                        text: "Current Tab"
                      },
                      {
                        key: "_blank",
                        text: "New Tab"
                      },
                      {
                        key: "_parent",
                        text: "Current Browser - New Window"
                      },
                      {
                        key: "_top",
                        text: "New Browser"
                      }                      
                    ],
                    required: true
                  }
                ],
                disabled: false
              })
            }                                      
            break;
          case 7:
            if(this.properties.groupTitle7 === undefined){this.properties.groupTitle7 = `Link Group ${x}`;}
            iconCode = PropertyFieldIconPicker('iconPicker7', {
              currentIcon: this.properties.iconPicker7,
              key: "iconPickerId",
              onSave: (icon: string) => { console.log(icon); this.properties.iconPicker7 = icon;},
              onChanged:(icon: string) => { console.log(icon);  },
              buttonLabel: "Select Icon",
              renderOption: "panel",
              properties: this.properties,
              onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
              label: "Icon Picker for Group "+x            
            });

            iconColour = PropertyFieldColorPicker('iconColour7', {
                        label: 'Icon Colour',
                        selectedColor: this.properties.iconColour7,
                        onPropertyChange: this.onPropertyPaneFieldChanged,
                        properties: this.properties,
                        disabled: false,
                        debounce: 1000,
                        isHidden: false,
                        alphaSliderHidden: false,
                        style: PropertyFieldColorPickerStyle.Full,
                        iconName: 'Precipitation',
                        key: 'colorFieldId'
                      });

            if(!this.properties.useList){ 
              linksCode = PropertyFieldCollectionData("linksGroup7", {
                key: "linksGroup7",
                label: "Group 7 Links",
                panelHeader: "Enter the links for Group "+x,
                manageBtnLabel: "Manage Links Group "+x,
                value: this.properties.linksGroup7,
                fields: [
                  {
                    id: "LinkTitle",
                    title: "Link Title",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkURL",
                    title: "Link URL",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkBrowse",
                    title: "Link Browse",
                    type: CustomCollectionFieldType.dropdown,
                    options: [
                      {
                        key: "_self",
                        text: "Current Tab"
                      },
                      {
                        key: "_blank",
                        text: "New Tab"
                      },
                      {
                        key: "_parent",
                        text: "Current Browser - New Window"
                      },
                      {
                        key: "_top",
                        text: "New Browser"
                      }                      
                    ],
                    required: true
                  }
                ],
                disabled: false
              })
            }                                      
            break;
          case 8:
            if(this.properties.groupTitle8 === undefined){this.properties.groupTitle8 = `Link Group ${x}`;}
            iconCode = PropertyFieldIconPicker('iconPicker8', {
              currentIcon: this.properties.iconPicker8,
              key: "iconPickerId",
              onSave: (icon: string) => { console.log(icon); this.properties.iconPicker8 = icon;},
              onChanged:(icon: string) => { console.log(icon);  },
              buttonLabel: "Select Icon",
              renderOption: "panel",
              properties: this.properties,
              onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
              label: "Icon Picker for Group "+x            
            });

            iconColour = PropertyFieldColorPicker('iconColour8', {
                        label: 'Icon Colour',
                        selectedColor: this.properties.iconColour8,
                        onPropertyChange: this.onPropertyPaneFieldChanged,
                        properties: this.properties,
                        disabled: false,
                        debounce: 1000,
                        isHidden: false,
                        alphaSliderHidden: false,
                        style: PropertyFieldColorPickerStyle.Full,
                        iconName: 'Precipitation',
                        key: 'colorFieldId'
                      });

            if(!this.properties.useList){ 
              linksCode = PropertyFieldCollectionData("linksGroup8", {
                key: "linksGroup8",
                label: "Group 8 Links",
                panelHeader: "Enter the links for Group "+x,
                manageBtnLabel: "Manage Links Group "+x,
                value: this.properties.linksGroup8,
                fields: [
                  {
                    id: "LinkTitle",
                    title: "Link Title",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkURL",
                    title: "Link URL",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkBrowse",
                    title: "Link Browse",
                    type: CustomCollectionFieldType.dropdown,
                    options: [
                      {
                        key: "_self",
                        text: "Current Tab"
                      },
                      {
                        key: "_blank",
                        text: "New Tab"
                      },
                      {
                        key: "_parent",
                        text: "Current Browser - New Window"
                      },
                      {
                        key: "_top",
                        text: "New Browser"
                      }                      
                    ],
                    required: true
                  }
                ],
                disabled: false
              })
            }                                      
            break;
          case 9:
            if(this.properties.groupTitle9 === undefined){this.properties.groupTitle9 = `Link Group ${x}`;}
            iconCode = PropertyFieldIconPicker('iconPicker9', {
              currentIcon: this.properties.iconPicker9,
              key: "iconPickerId",
              onSave: (icon: string) => { console.log(icon); this.properties.iconPicker9 = icon;},
              onChanged:(icon: string) => { console.log(icon);  },
              buttonLabel: "Select Icon",
              renderOption: "panel",
              properties: this.properties,
              onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
              label: "Icon Picker for Group "+x            
            });

            iconColour = PropertyFieldColorPicker('iconColour9', {
                        label: 'Icon Colour',
                        selectedColor: this.properties.iconColour9,
                        onPropertyChange: this.onPropertyPaneFieldChanged,
                        properties: this.properties,
                        disabled: false,
                        debounce: 1000,
                        isHidden: false,
                        alphaSliderHidden: false,
                        style: PropertyFieldColorPickerStyle.Full,
                        iconName: 'Precipitation',
                        key: 'colorFieldId'
                      });

            if(!this.properties.useList){ 
              linksCode = PropertyFieldCollectionData("linksGroup9", {
                key: "linksGroup9",
                label: "Group 9 Links",
                panelHeader: "Enter the links for Group "+x,
                manageBtnLabel: "Manage Links Group "+x,
                value: this.properties.linksGroup9,
                fields: [
                  {
                    id: "LinkTitle",
                    title: "Link Title",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkURL",
                    title: "Link URL",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkBrowse",
                    title: "Link Browse",
                    type: CustomCollectionFieldType.dropdown,
                    options: [
                      {
                        key: "_self",
                        text: "Current Tab"
                      },
                      {
                        key: "_blank",
                        text: "New Tab"
                      },
                      {
                        key: "_parent",
                        text: "Current Browser - New Window"
                      },
                      {
                        key: "_top",
                        text: "New Browser"
                      }                      
                    ],
                    required: true
                  }
                ],
                disabled: false
              })
            }                                      
            break;
          case 10:
            if(this.properties.groupTitle10 === undefined){this.properties.groupTitle10 = `Link Group ${x}`;}
            iconCode = PropertyFieldIconPicker('iconPicker10', {
              currentIcon: this.properties.iconPicker10,
              key: "iconPickerId",
              onSave: (icon: string) => { console.log(icon); this.properties.iconPicker10 = icon;},
              onChanged:(icon: string) => { console.log(icon);  },
              buttonLabel: "Select Icon",
              renderOption: "panel",
              properties: this.properties,
              onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
              label: "Icon Picker for Group "+x            
            });

            iconColour = PropertyFieldColorPicker('iconColour10', {
                        label: 'Icon Colour',
                        selectedColor: this.properties.iconColour10,
                        onPropertyChange: this.onPropertyPaneFieldChanged,
                        properties: this.properties,
                        disabled: false,
                        debounce: 1000,
                        isHidden: false,
                        alphaSliderHidden: false,
                        style: PropertyFieldColorPickerStyle.Full,
                        iconName: 'Precipitation',
                        key: 'colorFieldId'
                      });

            if(!this.properties.useList){ 
              linksCode = PropertyFieldCollectionData("linksGroup10", {
                key: "linksGroup10",
                label: "Group 10 Links",
                panelHeader: "Enter the links for Group "+x,
                manageBtnLabel: "Manage Links Group "+x,
                value: this.properties.linksGroup10,
                fields: [
                  {
                    id: "LinkTitle",
                    title: "Link Title",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkURL",
                    title: "Link URL",
                    type: CustomCollectionFieldType.string,
                    required: true
                  },
                  {
                    id: "LinkBrowse",
                    title: "Link Browse",
                    type: CustomCollectionFieldType.dropdown,
                    options: [
                      {
                        key: "_self",
                        text: "Current Tab"
                      },
                      {
                        key: "_blank",
                        text: "New Tab"
                      },
                      {
                        key: "_parent",
                        text: "Current Browser - New Window"
                      },
                      {
                        key: "_top",
                        text: "New Browser"
                      }                      
                    ],
                    required: true
                  }
                ],
                disabled: false
              })
            }                                      
            break;            
        }

        if(this.properties.useList){

            var singlePanel: IPropertyPaneGroup = {
              groupName: "Link Group "+x,
              groupFields: [
                PropertyPaneTextField('groupTitle'+x, {
                  label: `Link Group ${x} Name`,
                  value: `Link Group ${x}`,
                  placeholder: "Please enter the link group name"  //,"description": "Name property field"
                }),
                iconCode,
                iconColour 
              ],
              isCollapsed: true,
            };
            groupPanels.push(singlePanel);          
        }else{
          var singlePanel: IPropertyPaneGroup = {
            groupName: "Link Group "+x,
            groupFields: [
              PropertyPaneTextField('groupTitle'+x, {
                label: `Link Group ${x} Name`,
                value: `Link Group ${x}`,
                placeholder: "Please enter the link group name"  //,"description": "Name property field"
              }),
              iconCode,
              iconColour,
              linksCode
            ],
            isCollapsed: true,
          };
          groupPanels.push(singlePanel);
        }

        console.log("iconPicker1 name",this.properties.iconPicker1);

        linkList={};
      }  // end for loop for groups    

    }else{

      var singlePanel: IPropertyPaneGroup = {
        groupName: "Link Groups",
        groupFields: [
          PropertyPaneLabel('',{
            text:"Please choose the number of groups required from page 1"
            })
        ]
      }  
      groupPanels.push(singlePanel);

      var linksPanel: IPropertyPaneGroup = {
        groupName: "Links for Groups",
        groupFields: [
          PropertyPaneLabel('',{
            text:"Please choose the number of groups required from page 1"
          })    
        ]
      } 
      linkPanels.push(linksPanel); 
    }

    if(this.properties.useList){
      linkList=PropertyPaneButton('', {
        text: "Open List",
        buttonType: PropertyPaneButtonType.Primary,
        onClick: this.buttonClick.bind(this) 
      })
    }
    
    return {
      pages: [
        {
          header: {
            description: "Page 1 - App Setup"
          },
          groups: [
            {
              groupName: "App Details",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "App title to display"
                }),
                PropertyPaneSlider('numGroups', {
                  label:'How Many Link Groups? (max 10)',
                  min:0,
                  max:10,
                  value:0
                }),
                PropertyPaneCheckbox('useList', {
                  text: 'Use SharePoint List as link data?'
                }),
                linkList,
                PropertyPaneLabel('',{
                  text:"Please go to page 2 to setup the link groups"
                })  
              ]
            }
          ]
        },
        { //Page 2
          displayGroupsAsAccordion: true,
          header : {
            description : "Page 2 - Groups Setup"
          },
          groups : groupPanels
        },
        /*
        { //Page 3
          displayGroupsAsAccordion: true,
          header: {
            description: "Page 3 – Group Links"
          },
          groups: linkPanels
        }
        */
      ]
    };
  }
}