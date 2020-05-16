import * as React from 'react';
import * as strings from 'PlannerReportsWebPartStrings';
import styles from './PlannerReports.module.scss';
import { IPlannerReportsProps } from './IPlannerReportsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import { Button, Dropdown } from 'office-ui-fabric-react/lib';

export default class PlannerReports extends React.Component<IPlannerReportsProps, {}> {

 public onFormReportClick()
 {
    
 }

  public render(): React.ReactElement<IPlannerReportsProps> {
    return (
      <div className={ styles.plannerReports }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <Button name = "formReport" onClick={this.onFormReportClick} text = {strings.CreateReportButton}/>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
