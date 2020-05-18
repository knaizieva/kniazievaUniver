import * as React from 'react';
import { DefaultButton, PrimaryButton, Stack, IStackTokens, Label, FontSizes } from 'office-ui-fabric-react';
import * as strings from 'PlannerReportsWebPartStrings';
import styles from './PlannerReports.module.scss';
import { IPlannerReportsProps } from './IPlannerReportsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox, ICheckboxProps } from 'office-ui-fabric-react/lib/Checkbox';
import * as docx from "docx";
const { TextRun, Document, Packer, Paragraph, Table, TableCell, TableRow } = docx;

interface IPlannerReportsState {
  loading: boolean;
  selectedSegment?: IDropdownOption;
  segmentOptions?: IDropdownOption[];
  taskOptions?: IDropdownOption[];
  selectedTasks?: string[];
  onlyCompletedTaskChecked?: boolean;
  allTaskChecked?: boolean;
}
const dropdownStyles: Partial<IDropdownStyles> = {
  root: { display: "flex" },
  dropdown: { minWidth: 200 }, label: { paddingRight: "5px" }
};
export default class PlannerReports extends React.Component<IPlannerReportsProps, IPlannerReportsState> {
  private rawTasks: MicrosoftGraph.PlannerTask[];
  constructor(props) {
    super(props);
    this.onFormReportClick = this.onFormReportClick.bind(this);
    this.getBuckets = this.getBuckets.bind(this);
    this.onSegmentChange = this.onSegmentChange.bind(this);
    this.onTaskChange = this.onTaskChange.bind(this);
    this.onOnlyFinishedTaskChecked = this.onOnlyFinishedTaskChecked.bind(this);
    this.createAndSaveDocx = this.createAndSaveDocx.bind(this);
    this.state = { loading: false, allTaskChecked: true };
  }

  private async onFormReportClick() {
    if (this.props && this.props.plan) {
      this.setState({ loading: true });
      //let tasks: MicrosoftGraph.PlannerTask[] = await this.props.getTasks(this.props.plan, );

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
    this.rawTasks = await this.props.getTasks(this.props.plan);
    let taskOpts: IDropdownOption[] = this.rawTasks.map(t => { return { key: t.id, text: t.title }; });
    this.setState({ taskOptions: taskOpts });
  }

  private async getBucketTasks(segmentKey: string | number) {
    this.rawTasks = await this.props.getBucketTasks(segmentKey);
    let taskOpts: IDropdownOption[] = this.rawTasks
      .filter(v => {
        if (this.state.onlyCompletedTaskChecked) {
          return v.percentComplete == 100;
        }
        return true;
      })
      .map(t => { return { key: t.id, text: t.title }; });
    this.setState({ taskOptions: taskOpts });
  }

  private onTaskChange(event, option: IDropdownOption): void {
    if (option) {
      this.setState({
        selectedTasks: option.selected ?
          [...this.state.selectedTasks, option.key as string] :
          this.state.selectedTasks.filter(key => key !== option.key)
      });
    }
  }

  private onOnlyFinishedTaskChecked(event, checked: boolean): void {
    this.setState({ onlyCompletedTaskChecked: checked });
  }
  private onAllTaskChecked(event, checked: boolean):void {

  }

  public render(): React.ReactElement<IPlannerReportsProps> {
    let p = this.props;
    return (
      <div className={styles.plannerReports}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <DefaultButton onClick={this.createAndSaveDocx}
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
          <div className={styles.row}>
            <div className={styles.column}>
              <Checkbox label={strings.CompletedTaskLabel} checked={this.state.onlyCompletedTaskChecked}
                onChange={this.onOnlyFinishedTaskChecked} />
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
  private createAndSaveDocx(): void {

   // let tableRows: TableRow[] = [];




    const table = new Table({
      rows: []
  });

    const doc = new Document();
    let date: Date = new Date();
    let kievYear: string = `Київ ${date.getFullYear()}`;
    doc.addSection({
      properties: {},
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: "«НАЦІОНАЛЬНИЙ УНІВЕРСИТЕТ БІОРЕСУРСІВ ТА ПРИРОДОКОРИСТУВАННЯ УКРАЇНИ»",
              size: 28
            })
          ],
          alignment: docx.AlignmentType.CENTER
        }),
        new Paragraph({
          spacing: {
            before: 4000
          },
          children: [
            new TextRun({
              text: "Звіт",
              size: 32,
              bold: true
            })
          ],
          alignment: docx.AlignmentType.CENTER
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Факультету інформаційний технологій",
              size: 28
            })
          ],
          alignment: docx.AlignmentType.CENTER
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: "Про результати роботи відділів",
              size: 28
            })
          ],
          alignment: docx.AlignmentType.CENTER
        }),
        new Paragraph({
          spacing: {
            before: 6000
          },
          children: [
            new TextRun({
              text: kievYear,
              size: 28,
              bold: true
            })
          ],
          alignment: docx.AlignmentType.CENTER
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: this.state.selectedSegment.text,
              size: 28,
              bold: true
            })
          ],
          alignment: docx.AlignmentType.CENTER
        })

      ]
    });
    Packer.toBuffer(doc).then(buf => {
      this.props.saveFile(buf, buf.byteLength, `Test_${date.getDate()}_${date.getMonth()}_${date.getFullYear()}.docx`);
    });
  }
}
