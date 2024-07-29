import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AnniversaryWebPartWebPartStrings';
import AnniversaryWebPart from './components/AnniversaryWebPart';
import { IAnniversaryWebPartProps } from './components/IAnniversaryWebPartProps';

// Import the necessary classes from PnPJS
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/views";
import "@pnp/sp/fields";
import { IODataList, IODataView, IOField } from '../../interfaces';
import { getSP } from '../../pnpjs-config';

export interface IAnniversaryWebPartWebPartProps {
  description: string;
  speed: number;
  height: number;
  listId: string;
  viewId: string;
  nameFieldId: string;
  dateFieldId: string;
  dateFormat: string;
  milestoneFieldId: string;
}

export default class AnniversaryWebPartWebPart extends BaseClientSideWebPart<IAnniversaryWebPartWebPartProps> {

  private ddlSharePointLists: IPropertyPaneDropdownOption [] =  [];
  private ddlSharePointViews: IPropertyPaneDropdownOption [] =  [];
  private ddlSharePointFields: IPropertyPaneDropdownOption [] = [];
  private ddlSharePointUserFields: IPropertyPaneDropdownOption [] = [];
  
  private _sp: SPFI;

  public render(): void {
    const element: React.ReactElement<IAnniversaryWebPartProps> = React.createElement(
      AnniversaryWebPart,
      {
        description: this.properties.description,
        speed: this.properties.speed,
        height: this.properties.height,
        listId: this.properties.listId,
        viewId: this.properties.viewId,
        nameFieldId: this.properties.nameFieldId,
        dateFieldId: this.properties.dateFieldId,
        dateFormat: this.properties.dateFormat,
        milestoneFieldId: this.properties.milestoneFieldId
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {

    await super.onInit();
    
    //Initialize our _sp object that we can then use in other packages without having to pass around the context.
    // Check out pnpjsConfig.ts for an example of a project setup file.
    this._sp = getSP(this.context);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    
    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void) => {
      
      this._sp.web.lists().then((lists: IODataList[]): void =>
      {
        const listOptions: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        
        lists.map((list: IODataList) => {
          if (list.BaseTemplate === 100)       // Custom list
            listOptions.push( { key: list.Id, text: list.Title });
        });

        resolve(listOptions);
      })
      .catch(err => console.error(err));

    })

  }

  private loadViews(): Promise<IPropertyPaneDropdownOption[]> {
    
    if (!this.properties.listId) {
      return Promise.resolve(this.ddlSharePointViews);
    }

    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void) => {
      
      this._sp.web.lists.getById(this.properties.listId).views().then((views:IODataView[]): void => 
      {
        const viewOptions: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        
        views.map((view: IODataView) => {
          viewOptions.push( { key: view.Id, text: view.Title });
        });

        resolve(viewOptions);
      })
      .catch(err => console.error(err));

    });

  }

  private loadListFieldNames(filterUserFieldsOnly: boolean = false): Promise<IPropertyPaneDropdownOption[]> {

    const fieldOptions: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();

    if (!this.properties.listId || !this.properties.viewId) {
      return Promise.resolve(this.ddlSharePointFields);
    }

    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void) => {
      
      this._sp.web.lists.getById(this.properties.listId).fields().then((fields:IOField[]): void => {
  
        fields.map((field: IOField) => {

          if (!field.Hidden && !field.ReadOnlyField && (!filterUserFieldsOnly || field.TypeAsString === "User"))
            fieldOptions.push( { key: field.InternalName, text: field.Title });  

        });

        resolve(fieldOptions);

      })
      .catch(err => console.error(err));

    });
  }

  protected onPropertyPaneConfigurationStart(): void
  {
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'SharePoint lists');

    this.loadLists().then((listOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
      this.ddlSharePointLists = listOptions;
      this.context.propertyPane.refresh();
      return this.loadViews();
    }).then((viewOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
      this.ddlSharePointViews = viewOptions;
      this.context.propertyPane.refresh();
      return this.loadListFieldNames();
    }).then((fieldOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
      this.ddlSharePointFields = fieldOptions;
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      return this.loadListFieldNames(true);
    }).then((fieldOptions: IPropertyPaneDropdownOption[]): void => {
      this.ddlSharePointUserFields = fieldOptions;
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render();
    })
    .catch(err => console.error(err));
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string): void {
    if (propertyPath === 'listId' && newValue) {
      // push new value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();

      // communicate loading items
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'list views');

      this.loadViews()
        .then((listOptions: IPropertyPaneDropdownOption[]): void => {
          // store field names
          this.ddlSharePointViews = listOptions;

          // Default to the first view
          this.properties.viewId = listOptions[0].key.toString();
          this.onPropertyPaneFieldChanged('viewId', '', this.properties.viewId);

          // clear status indicator
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          // re-render the web part as clearing the loading indicator removes the web part body
          this.render();
          // refresh the item selector control by repainting the property pane
          this.context.propertyPane.refresh();
        })
        .catch(err => console.error(err));
    }
    else if (propertyPath === 'viewId' && newValue) {
      // push new value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      // get previously selected library fields
      const previousField1: string = this.properties.nameFieldId;
      const previousField2: string = this.properties.dateFieldId;
      const previousField3: string = this.properties.milestoneFieldId;

      // reset selected fields
      this.properties.nameFieldId = "-- Select One --";
      this.properties.dateFieldId = "-- Select One --";
      this.properties.milestoneFieldId = "-- Select One --";

      // push new field values
      this.onPropertyPaneFieldChanged('nameFieldId', previousField1, this.properties.nameFieldId);
      this.onPropertyPaneFieldChanged('dateFieldId', previousField2, this.properties.dateFieldId);
      this.onPropertyPaneFieldChanged('milestoneFieldId', previousField3, this.properties.milestoneFieldId);

      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // communicate loading items
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'list fields');

      this.loadListFieldNames()
        .then((listOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
          // store field names
          this.ddlSharePointFields = listOptions;
          // clear status indicator
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          
          return this.loadListFieldNames(true);
        })        
        .then((listOptions: IPropertyPaneDropdownOption[]): void => {
          // store field names
          this.ddlSharePointUserFields = listOptions;
          // clear status indicator
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          // re-render the web part as clearing the loading indicator removes the web part body
          this.render();
          // refresh the item selector control by repainting the property pane
          this.context.propertyPane.refresh();
        })
        .catch(err => console.error(err));
        
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneSlider('speed', {
                  label: strings.SpeedFieldLabel,
                  min: 2,
                  max: 100,
                  step: 2,
                  showValue: true
                }),
                PropertyPaneTextField('height', {
                  label: strings.HeightFieldLabel
                })
              ]
            },
            {
              groupName: strings.SharePointGroupName,
              groupFields: [
                PropertyPaneDropdown('listId', {
                  label: strings.SharePointListLabel,
                  options: this.ddlSharePointLists
                }),
                PropertyPaneDropdown('viewId', {
                  label: strings.SharePointViewLabel,
                  options: this.ddlSharePointViews
                }),
                PropertyPaneDropdown('nameFieldId', {
                  label: strings.SharePointNameFieldLabel,
                  options: this.ddlSharePointUserFields
                }),
                PropertyPaneDropdown('dateFieldId', {
                  label: strings.SharePointDateFieldLabel,
                  options: this.ddlSharePointFields
                }),
                PropertyPaneTextField('dateFormat', {
                  label: strings.SharePointDateFormatLabel
                }),
                PropertyPaneDropdown('milestoneFieldId', {
                  label: strings.SharePointMilestoneFieldLabel,
                  options: this.ddlSharePointFields
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
