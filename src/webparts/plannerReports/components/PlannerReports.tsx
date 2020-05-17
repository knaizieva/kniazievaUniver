import * as React from 'react';
import { DefaultButton, PrimaryButton, Stack, IStackTokens, Label } from 'office-ui-fabric-react';
import * as strings from 'PlannerReportsWebPartStrings';
import styles from './PlannerReports.module.scss';
import { IPlannerReportsProps } from './IPlannerReportsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';

interface IPlannerReportsState {
  loading: boolean;
  selectedSegment?: IDropdownOption;
  segmentOptions?: IDropdownOption[];
  taskOptions?: IDropdownOption[];
  selectedTasks?: string[];
}
const dropdownStyles: Partial<IDropdownStyles> = {
  root: { display: "flex" },
  dropdown: { minWidth: 200, maxWidth: 300 }, label: { paddingRight: "5px" }
};
export default class PlannerReports extends React.Component<IPlannerReportsProps, IPlannerReportsState> {
  constructor(props) {
    super(props);
    this.onFormReportClick = this.onFormReportClick.bind(this);
    this.getBuckets = this.getBuckets.bind(this);
    this.onSegmentChange = this.onSegmentChange.bind(this);
    this.onTaskChange = this.onTaskChange.bind(this);
    this.state = { loading: false };
  }

  private async onFormReportClick() {
    if (this.props && this.props.plan) {
      this.setState({ loading: true });
      let tasks: MicrosoftGraph.PlannerTask[] = await this.props.getTasks(this.props.plan);

      this.setState({ loading: false });
    }
  }

  private async getBuckets() {
    if (this.props && this.props.plan) {
      this.setState({ loading: true });
      let buckets: MicrosoftGraph.PlannerBucket[] = await this.props.getBuckets(this.props.plan);
      let segments: IDropdownOption[] = buckets.map(v => { return { key: v.id, text: v.name }; });
      this.setState({ segmentOptions: segments });
      this.setState({ loading: false, selectedTasks: [] });
    }
  }



  private onSegmentChange(event, option: IDropdownOption) {
    this.setState({ selectedSegment: option });
    this.getBucketTasks(option.key);
  }

  private async getTasks(segmentKey: string | number) {
    let tasks: MicrosoftGraph.PlannerTask[] = await this.props.getTasks(this.props.plan);
    let taskOpts: IDropdownOption[] = tasks.map(t => { return { key: t.id, text: t.title }; });
    this.setState({ taskOptions: taskOpts });
  }

  private async getBucketTasks(segmentKey: string | number) {
    let tasks: MicrosoftGraph.PlannerTask[] = await this.props.getBucketTasks(segmentKey);
    let taskOpts: IDropdownOption[] = tasks.map(t => { return { key: t.id, text: t.title }; });
    this.setState({ taskOptions: taskOpts });
  }

  private onTaskChange(event, option: IDropdownOption):void {
    if (option) {
      this.setState({selectedTasks: option.selected ? 
        [...this.state.selectedTasks, option.key as string] : 
        this.state.selectedTasks.filter(key => key !== option.key)});
    }
  }


  public render(): React.ReactElement<IPlannerReportsProps> {
    let p = this.props;
    return (
      <div className={styles.plannerReports}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <DefaultButton name="formReport" onClick={this.getBuckets}
                text={this.state.loading ? strings.Loading + "..." : strings.CreateReportButton} />
            </div>
            <div className={styles.column}>
              <Dropdown
                label={strings.SegmentLabel}
                selectedKey={this.state.selectedSegment ? this.state.selectedSegment.key : undefined}
                onChange={this.onSegmentChange}
                placeholder=""
                options={this.state.segmentOptions}
                styles={dropdownStyles}
              />
            </div>
            <div className={styles.column}>
              <Dropdown
                label={strings.TasksLabel}
                selectedKeys={this.state.selectedTasks}
                onChange={this.onTaskChange}
                placeholder=""
                multiSelect
                options={this.state.taskOptions}
                styles={dropdownStyles}
              />
            </div>
          </div>
        </div>
      </div>
    );
  }
  public componentDidMount(): void {
    if (this.props.plan) {
      this.getBuckets();
    }
  }
  public componentDidUpdate(prevProps: IPlannerReportsProps): void {
    if (prevProps.plan != this.props.plan && this.props.plan) {
      this.getBuckets();
    }
  }
}
