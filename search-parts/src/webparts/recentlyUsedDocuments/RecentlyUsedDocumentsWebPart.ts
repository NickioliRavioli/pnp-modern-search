import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-property-pane";
import * as strings from 'RecentlyUsedDocumentsWebPartStrings';
import RecentlyUsedDocuments from './components/RecentlyUsedDocuments';
import { IRecentlyUsedDocumentsProps } from './components/IRecentlyUsedDocumentsProps';

export interface IRecentlyUsedDocumentsWebPartProps {
  title: string;
  nrOfItems: number;
  siteFilter: string;
}

export default class RecentlyUsedDocumentsWebPart extends BaseClientSideWebPart<IRecentlyUsedDocumentsWebPartProps> {
  private graphClient: any;
  private propertyFieldNumber : any;

  public onInit(): Promise<void> {
    this.initializeProperties();
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.msGraphClientFactory
        .getClient('3')
        .then((client) => {
          this.graphClient = client;
          resolve();
        }, err => reject(err));
    });
  }

  private initializeProperties() {
    this.properties.title = this.properties.title ? this.properties.title : "";
    this.properties.nrOfItems = this.properties.nrOfItems ? this.properties.nrOfItems : 10;
    this.properties.siteFilter = this.properties.siteFilter ? this.properties.siteFilter : "https://dbctcomau.sharepoint.com/sites/CDC";
  }


  public render(): void {
    const element: React.ReactElement<IRecentlyUsedDocumentsProps> = React.createElement(
      RecentlyUsedDocuments,
      {
        title: this.properties.title,
        nrOfItems: this.properties.nrOfItems,
        siteFilter: this.properties.siteFilter,
        context: this.context,
        graphClient: this.graphClient,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }
  
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //executes only before property pane is loaded.
  protected async loadPropertyPaneResources(): Promise<void> {
    // import additional controls/components

    const { PropertyFieldNumber} = await import(
      /* webpackChunkName: 'pnp-propcontrols-number' */
      '@pnp/spfx-property-controls/lib/propertyFields/number'
    );

    this.propertyFieldNumber = PropertyFieldNumber;
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
              groupFields: [
                this.propertyFieldNumber("nrOfItems", {
                  key: "nrOfItems",
                  label: strings.NrOfDocumentsToShow,
                  value: this.properties.nrOfItems,
                  minValue: 1,
                  maxValue: 20
                }),
                PropertyPaneTextField("siteFilter", {
                  label: "Site filter",
                  multiline: false,
                  placeholder: 'e.g. https://dbctcomau.sharepoint.com/sites/CDC',
                  value: this.properties.siteFilter
                })
              ]
            }
          ]
        }
      ]
    };
  }
}