import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IWebPartContext
} from '@microsoft/sp-webpart-base';
import { IDigestCache, DigestCache } from '@microsoft/sp-http';
import * as strings from 'FileUploadWebPartStrings';
import FileUpload from './components/FileUpload';
import { IFileUploadProps } from './components/IFileUploadProps';
import * as loader from '@microsoft/sp-loader';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '../../PropertyFieldListPicker';
import { PropertyPaneDropdown } from '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneDropdown/PropertyPaneDropdown';
import { PropertyPaneAsyncDropdown } from '../../controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown';
import { SiteQueryService } from '../../services/SiteQueryService';
import { ISiteQueryService } from '../../services/ISiteQueryService';
import { IDropdownOption } from 'office-ui-fabric-react';
import { update, get, isEmpty } from '@microsoft/sp-lodash-subset';
import { SiteQueryConstants } from '../../../lib/common/constants/SiteQueryConstants';

export interface IFileUploadWebPartProps {
  listName:string;
  fileTypes:string;
  queryString:string;
  uploadFilesTo:string;
  siteUrl:string;
  webUrl:string;
}
require("./filepicker.css");
require("./dropzone.css");

export default class FileUploadWebPart extends BaseClientSideWebPart<IFileUploadWebPartProps> {
  private SiteQueryService: ISiteQueryService;
  private siteUrlDropdown: PropertyPaneAsyncDropdown;
  private webUrlDropdown: PropertyPaneAsyncDropdown;
  private filesToDropdown: any;
  private listTitleDropdown: any;
  private fileTypeTextField: any;
  private queryStringTextField: any;
  public digest:string="";
  public constructor(context:IWebPartContext){
    super();    
    loader.SPComponentLoader.loadCss('https://code.ionicframework.com/ionicons/2.0.1/css/ionicons.min.css');
  }
  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.SiteQueryService = new SiteQueryService(this.context, this.context.spHttpClient);
      const digestCache: IDigestCache = this.context.serviceScope.consume(DigestCache.serviceKey);
      digestCache.fetchDigest(this.context.pageContext.web.serverRelativeUrl).then((digest: string): void => {
        // use the digest here
        this.digest=digest;
        resolve();
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IFileUploadProps > = React.createElement(
      FileUpload,
      {
        digest:this.digest,
        context:this.context,
        listName:this.properties.listName,
        fileTypes:this.properties.fileTypes,
        queryString:this.properties.queryString,
        uploadFilesTo:this.properties.uploadFilesTo,
        siteUrl:this.properties.siteUrl,
        webUrl:this.properties.webUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let firstCascadingLevelDisabled = !this.properties.siteUrl;
    let secondCascadingLevelDisabled = !this.properties.siteUrl || !this.properties.webUrl;

        // Create SiteCollection Dropdown
        this.siteUrlDropdown = new PropertyPaneAsyncDropdown(SiteQueryConstants.propertySiteUrl, {
          label: strings.SiteUrlFieldLabel,
          loadOptions: this.loadSiteUrlOptions.bind(this),
          loadingLabel: strings.LoadingSiteDropdown,
          errorLabelFormat : strings.ErrorOnLoadingSiteDropdown,
          onPropertyChange: this.onCustomPropertyPaneChange.bind(this),
          selectedKey: this.properties.siteUrl
        });
               // Create Web (subsite) Dropdown  
    this.webUrlDropdown = new PropertyPaneAsyncDropdown(SiteQueryConstants.propertyWebUrl, {
      label: strings.WebUrlFieldLabel,
      loadOptions: this.loadWebUrlOptions.bind(this) ,
      loadingLabel: strings.LoadingWebDropdown,
      errorLabelFormat : strings.ErrorOnLoadingWebDropdown,
      onPropertyChange: this.onCustomPropertyPaneChange.bind(this),
      selectedKey: this.properties.webUrl,
      disabled: firstCascadingLevelDisabled
    });
    this.listTitleDropdown = PropertyFieldListPicker('listName', {
      label: 'Select a list or library',
      selectedList: this.properties.listName,
      includeHidden: false,
      //baseTemplate: 109,
      orderBy: PropertyFieldListPickerOrderBy.Title,
      // multiSelect: false,
      disabled: secondCascadingLevelDisabled,
      onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
      properties: this.properties,
      context: this.context,                  
      onGetErrorMessage: null,
      deferredValidationTime: 0,
      key: 'listPickerFieldId',
      webUrl: this.properties.webUrl || this.context.pageContext.web.absoluteUrl      
    });
    this.filesToDropdown = PropertyPaneDropdown('uploadFilesTo',{
      label:'Upload files to',
      options:[{key:'DocumentLibrary',text:'Document Library'},
               {key:'List',text:'As item attachments'} ]
    });
    this.fileTypeTextField = PropertyPaneTextField('fileTypes',{
      label:'File Types (use , as seperator)',                  
    });
    this.queryStringTextField = PropertyPaneTextField('queryString',{
      label:'Query String parameter',
      description:'If you want to attach files to a list item you need to define the ID of the item in a query string parameter, example: ID=1'
    });
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                this.filesToDropdown,
                this.siteUrlDropdown,
                this.webUrlDropdown,
                this.listTitleDropdown,
                this.fileTypeTextField,
                this.queryStringTextField
              ]
            }
          ]
        }
      ]
    };
  }
  private loadSiteUrlOptions(): Promise<IDropdownOption[]> {
    return this.SiteQueryService.getSiteUrlOptions();
  }
  private loadWebUrlOptions(): Promise<IDropdownOption[]> {
    return this.SiteQueryService.getWebUrlOptions(this.properties.siteUrl);
  }
  private onCustomPropertyPaneChange(propertyPath: string, newValue: any): void {
    const oldValue = get(this.properties, propertyPath);
    
    // Stores the new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    // Resets dependent property panes if needed
    this.resetDependentPropertyPanes(propertyPath);
    this.render();
    // Refreshes the web part manually because custom fields don't update since sp-webpart-base@1.1.1
    // https://github.com/SharePoint/sp-dev-docs/issues/594
    if (!this.disableReactivePropertyChanges)
      this.render();
      this.context.propertyPane.refresh(); 
  }
  private resetDependentPropertyPanes(propertyPath: string): void {
    if(propertyPath == SiteQueryConstants.propertySiteUrl) {
      this.resetWebUrlPropertyPane();
     
    }
 else if(propertyPath == SiteQueryConstants.propertyWebUrl) {
    this.resetListTitlePropertyPane();
   }
  }
  private resetWebUrlPropertyPane() {
    this.properties.webUrl = "";
    this.SiteQueryService.clearCachedWebUrlOptions();
    update(this.properties, SiteQueryConstants.propertyWebUrl, (): any => { return this.properties.webUrl; });
    this.webUrlDropdown.properties.selectedKey = "";
    this.webUrlDropdown.properties.disabled = isEmpty(this.properties.siteUrl);
    this.webUrlDropdown.render();
    
  }
  private resetListTitlePropertyPane() {
    this.context.propertyPane.refresh(); 
  }
}
