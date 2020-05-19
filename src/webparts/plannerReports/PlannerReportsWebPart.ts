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
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { ThemeSettingName, List } from 'office-ui-fabric-react';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { IListInfo } from '@pnp/sp/lists';
import { _Folder, IFolderInfo } from '@pnp/sp/folders/types';
import { IFileAddResult } from '@pnp/sp/files';

export interface IPlannerReportsWebPartProps {
  description: string;
  group: string;
  plan: string;
  library: string;
}

export default class PlannerReportsWebPart extends BaseClientSideWebPart<IPlannerReportsWebPartProps> {
  private groupDropdownOption: IPropertyPaneDropdownOption[];
  private planDropdownOption: IPropertyPaneDropdownOption[];
  private libraryDropdownOption: IPropertyPaneDropdownOption[];
  private groupDropdownDisabled: boolean = false;
  private planDropdownDisabled: boolean = false;
  private libraryDropdownDisabled: boolean = false;
  private loading: boolean = false;

  constructor()
  {
    super();
    this.getTasks = this.getTasks.bind(this);
    this.getTaskDetails = this.getTaskDetails.bind(this);
    this.getBuckets = this.getBuckets.bind(this);
    this.getBucketTasks = this.getBucketTasks.bind(this);
    this.saveFile = this.saveFile.bind(this);
    this.showPropertyPane = this.showPropertyPane.bind(this);
  }

  public render(): void {
    const element: React.ReactElement<IPlannerReportsProps> = React.createElement(
      PlannerReports,
      {
        description: this.properties.description,
        group: this.properties.group,
        plan: this.properties.plan,
        library: this.properties.library,
        getTasks: this.getTasks,
        getTaskDetails: this.getTaskDetails,
        getBuckets: this.getBuckets,
        getBucketTasks: this.getBucketTasks,
        saveFile: this.saveFile,
        showPropertyPane: this.showPropertyPane
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected getBuckets(planId:string): Promise<MicrosoftGraph.PlannerBucket[]>
  {
    return new Promise<MicrosoftGraph.PlannerBucket[]>((resolve, reject) => {
      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client.api(`planner/plans/${planId}/buckets`).get()
          .then(resp => {
            resolve(resp.value);
          })
          .catch(err => {
            reject(err);
          });
        });
    });
  }
  private showPropertyPane(): void {
    if(!this.context.propertyPane.isPropertyPaneOpen()) {
      this.context.propertyPane.open();
    }
  }
  private saveFile(file: any, size: number, fileName:string): Promise<IFileAddResult>
  {
    return new Promise<IFileAddResult>((resolve, reject) => {
      sp.web.lists.getById(this.properties.library).rootFolder.files
      .add(fileName, file, true)
      .then(r => {
        resolve(r);
      }).catch(e => {
        reject(e);
      });
    });
  }

  protected getBucketTasks(bucketId: string | number): Promise<MicrosoftGraph.PlannerTask[]> {
    return new Promise<MicrosoftGraph.PlannerTask[]>((resolve, reject) => {
      if (!bucketId) {
        return resolve([]);
      }
      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client.api(`planner/buckets/${bucketId}/tasks`)
          .get()
          .then(resp => {
            resolve(resp.value as MicrosoftGraph.PlannerTask[]);
          }).catch(err => {
            reject(err);
          });
        });
    });
  }

  protected getTasks(planId: string): Promise<MicrosoftGraph.PlannerTask[]> {
    
    return new Promise<MicrosoftGraph.PlannerTask[]>((resolve, reject) => {
      if (!planId) {
        return resolve([]);
      }
      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client.api(`planner/plans/${planId}/tasks`).get()
          .then(resp => {
            resolve(resp.value as MicrosoftGraph.PlannerTask[]);
          }).catch(err => {
            reject(err);
          });
        });
    });
  }

  private getTaskDetails(taskId:string): Promise<MicrosoftGraph.PlannerTaskDetails>
  {
    return new Promise<MicrosoftGraph.PlannerTaskDetails>((resolve, reject) => {
      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client.api(`planner/tasks/${taskId}/details`).get()
          .then(resp => {
            let r = resp;
          })
          .catch(err => {
            reject(err);
          });
    });
  });
}

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected async onInit(): Promise<void> {
    sp.setup(this.context);
    await super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //Get groups where current user is member of

  private async getGroups(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>((resolve, reject) => {
      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client
            .api('me/getMemberGroups')
            .post({ "securityEnabledOnly": false })
            .then(resp => {
              if (!resp.value.length) {
                resolve([]);
                return;
              }
              let queryParams: string = '';
              for (let i = 0; resp.value.length > i; i++) {
                if (i == 0) {
                  queryParams += `Id eq '${resp.value[i]}'`;
                } else {
                  queryParams += ` or Id eq '${resp.value[i]}'`;
                }
              }
              client.api(`groups?$filter=${queryParams}`).get()
                .then((groupResp) => {
                  resolve(groupResp.value.map((g: MicrosoftGraph.Group) => {
                    return { key: g.id, text: g.displayName };
                  }));
                });
            }).catch(err => {
              reject(err);
            });
        });
    });
  }

  private getPlans(groupId: string): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>((resolve, reject) => {
      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client.api(`groups/${groupId}/planner/plans`).get()
            .then(resp => {
              let plans: microsoftgraph.PlannerPlan[] = resp.value;
              resolve(
                plans.map(p => {
                  return { key: p.id, text: p.title };
                })
              );
            })
            .catch(err => {
              reject(err);
            });
        });
    });
  }

  private getLibs(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>((resolve, reject) => {
      sp.web.lists.filter("BaseTemplate eq 101")
        .get().then((resp: IListInfo[]) => {
          resolve(
            resp.map(l => {
              return { key: l.Id, text: l.Title };
            })
          );
        }).catch(err => {
          reject(err);
        });
    });
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (!newValue) {
      return;
    }
    this.loading = true;
    this.context.propertyPane.refresh();
    switch (propertyPath) {
      case 'group':
        this.planDropdownOption = await this.getPlans(newValue);
        this.planDropdownDisabled = false;
        break;
    }
    this.loading = false;
    this.context.propertyPane.refresh();
  }


  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    this.groupDropdownDisabled = !this.groupDropdownOption;
    this.libraryDropdownDisabled = !this.libraryDropdownOption;
    this.planDropdownDisabled =  !this.properties || !this.properties.group;

    if (this.groupDropdownOption) {
      return;
    }

    this.loading = true;
    this.context.propertyPane.refresh();
    this.groupDropdownOption = await this.getGroups();
    this.libraryDropdownOption = await this.getLibs();
    this.libraryDropdownDisabled = false;
    this.loading = false;
    this.groupDropdownDisabled = false;
    this.context.propertyPane.refresh();
  }

  //--------------

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                  disabled: this.groupDropdownDisabled,
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
