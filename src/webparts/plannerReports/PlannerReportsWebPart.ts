import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneLabel,
  PropertyPaneLink,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { 
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import * as strings from 'PlannerReportsWebPartStrings';
import PlannerReports from './components/PlannerReports';
import { IPlannerReportsProps } from './components/IPlannerReportsProps';
import { sp } from '@pnp/sp';
import '@pnp/sp/lists';
import "@pnp/sp/webs";
import { ThemeSettingName, List } from 'office-ui-fabric-react';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { IListInfo } from '@pnp/sp/lists';

export interface IPlannerReportsWebPartProps {
  description: string;
  group:string;
  plan:string;
  library:string;
}

export default class PlannerReportsWebPart extends BaseClientSideWebPart <IPlannerReportsWebPartProps> {
  private groupDropdownOption: IPropertyPaneDropdownOption[];
  private planDropdownOption: IPropertyPaneDropdownOption[];
  private libraryDropdownOption: IPropertyPaneDropdownOption[];
  private groupDropdownDisabled: boolean;
  private planDropdownDisabled: boolean;
  private libraryDropdownDisabled:boolean;
  private loading: boolean = false;
  public render(): void {
    const element: React.ReactElement<IPlannerReportsProps> = React.createElement(
      PlannerReports,
      {
        description: this.properties.description,
        group: this.properties.group,
        plan: this.properties.plan,
        library: this.properties.library
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected async onInit(): Promise<void>
  {
    sp.setup(this.context);
    await super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //Get groups where current user is member of

  private async getGroups(): Promise<IPropertyPaneDropdownOption[]>
  {
    return new Promise<IPropertyPaneDropdownOption[]>((resolve, reject) => {
    this.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient): void => {
      client
        .api('me/getMemberGroups')
        .post({"securityEnabledOnly": false})
        .then(resp => {
          if(!resp.value.length){
            resolve([]);
            return;
          }
          let queryParams:string = '';
          for(let i = 0; resp.value.length > i; i++)
          {
            if(i == 0)
            {
              queryParams += `Id eq '${resp.value[i]}'`;
            } else {
              queryParams +=  ` or Id eq '${resp.value[i]}'`;
            }
          }
          client.api(`groups?$filter=${queryParams}`).get()
          .then((groupResp) => {
            resolve(groupResp.value.map((g:MicrosoftGraph.Group) => {
              return {key: g.id, text: g.displayName};
            }));
          });
        }).catch(err => {
          reject(err);
        });
    });
    });
  }

  private getPlans(groupId: string):Promise<IPropertyPaneDropdownOption[]>
  {
    return new Promise<IPropertyPaneDropdownOption[]>((resolve, reject) =>{
      this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client.api(`groups/${groupId}/planner/plans`).get()
        .then(resp => {
          let plans: microsoftgraph.PlannerPlan[] = resp.value;
          resolve(
            plans.map(p => {
              return {key: p.id, text: p.title};
            })
          );
        })
        .catch(err =>{
          reject(err);
        });
      });
    });
  }

  private getLibs(): Promise<IPropertyPaneDropdownOption[]>
  {
    return new Promise<IPropertyPaneDropdownOption[]>((resolve, reject) => {
      sp.web.lists.filter("BaseTemplate eq 101")
      .get().then((resp: IListInfo[]) => {
        resolve(
          resp.map(l => {
            return {key: l.Id, text: l.Title};
          })
        );
      }).catch(err => {
        reject(err);
      });
    });
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if(!newValue)
    {
      return;
    }
    this.loading = true;
    this.context.propertyPane.refresh();
    switch (propertyPath) {
      case 'group':
        this.planDropdownOption = await this.getPlans(newValue);
        this.planDropdownDisabled = false;
        break;
      case 'plan':
        this.libraryDropdownOption = await this.getLibs();
        this.libraryDropdownDisabled = false;
        break;
    }
    this.loading = false;
    this.context.propertyPane.refresh();
  }


  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    this.groupDropdownDisabled = !this.groupDropdownDisabled;
    this.planDropdownDisabled = !this.properties.group;
    
    if(this.groupDropdownOption){
      return;
    }

    this.loading = true;
    this.groupDropdownOption = await this.getGroups();
    this.loading = false;
    this.groupDropdownDisabled = false;
    this.context.propertyPane.refresh(); 
  }

  //--------------

  protected  getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      showLoadingIndicator: this.loading,
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('group', {
                  disabled: !this.groupDropdownOption,
                  label: strings.GroupFieldLabel,
                  options: this.groupDropdownOption,
                }),
                PropertyPaneDropdown('plan', {
                  options: this.planDropdownOption,
                  label: strings.PlanFieldLabel, 
                  disabled: this.planDropdownDisabled
                }),
                PropertyPaneDropdown('library', {
                  options: this.libraryDropdownOption,
                  label: strings.LibraryFieldLabel,
                  disabled: this.libraryDropdownDisabled
                  })
              ]
            }
          ]
        }
      ]
    };
  }
}
