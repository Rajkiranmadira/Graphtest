import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { sp } from '@pnp/sp/presets/all';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';


import * as strings from 'GraphtestWebPartStrings';
import Graphtest from './components/Graphtest';
import { IGraphtestProps } from './components/IGraphtestProps';

export interface IGraphtestWebPartProps {
  description: string;
}

export default class GraphtestWebPart extends BaseClientSideWebPart<IGraphtestWebPartProps> {



  public render(): void {
    const element: React.ReactElement<IGraphtestProps> = React.createElement(
      Graphtest,
      {
        description: this.properties.description,
        context:this.context,
        siteUrl:this.context.pageContext.web.absoluteUrl
        
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(() => {
      sp.setup({
        ie11: true,  
        spfxContext:this.context as any
      });
    });
  }



 

  

  

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
