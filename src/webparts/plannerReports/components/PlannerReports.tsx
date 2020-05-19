import * as React from 'react';
import {
  DefaultButton, PrimaryButton, Stack, IStackTokens, MessageBar,
  MessageBarType, Label, FontSizes, MessageBarButton, ProgressIndicator, Link
} from 'office-ui-fabric-react';
import * as strings from 'PlannerReportsWebPartStrings';
import styles from './PlannerReports.module.scss';
import { IPlannerReportsProps } from './IPlannerReportsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox, ICheckboxProps } from 'office-ui-fabric-react/lib/Checkbox';
import * as docx from "docx";
import { TextRun, Document, Packer, Paragraph, Table, TableCell, TableRow } from "docx";

interface IPlannerReportsState {
  loading: boolean;
  selectedSegment?: IDropdownOption;
  segmentOptions?: IDropdownOption[];
  taskOptions?: IDropdownOption[];
  selectedTasks?: string[];
  onlyCompletedTaskChecked?: boolean;
  allTaskChecked?: boolean;
  erroMessage?: string;
  successfulMessage?: string;
  fileLink?: string;
}
const dropdownStyles: Partial<IDropdownStyles> = {
  root: { display: "flex" },
  dropdown: { minWidth: 200, maxWidth: 400 }, label: { paddingRight: "5px" }
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
    this.onSetPropertiesClick = this.onSetPropertiesClick.bind(this);
    this.onErrorMessDismiss = this.onErrorMessDismiss.bind(this);
    this.onSuccessMessDismiss = this.onSuccessMessDismiss.bind(this);
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

    let taskOpts: IDropdownOption[] = this.rawTasks
      .filter(v => {
        if (checked) {
          return v.percentComplete == 100;
        }
        return true;
      })
      .map(t => { return { key: t.id, text: t.title }; });
    this.setState({ taskOptions: taskOpts });

  }
  private onAllTaskChecked(event, checked: boolean): void {

  }

  public render(): React.ReactElement<IPlannerReportsProps> {
    let p = this.props;
    if (!this.props.group || !this.props.library || !this.props.plan) {
      return this.getSetPropertiesMessage();
    } else {
      return this.getBody();
    }

  }

  private onSetPropertiesClick(): void {
    this.props.showPropertyPane();
  }

  public getSetPropertiesMessage(): JSX.Element {
    return (
      <div className={styles.plannerReports}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <MessageBar
                messageBarType={MessageBarType.warning}
                isMultiline={false}
                //onDismiss={p.resetChoice}
                dismissButtonAriaLabel="Close"
                actions={
                  <div>
                    <MessageBarButton onClick={this.onSetPropertiesClick}>{strings.SetPropertiesLabel}</MessageBarButton>
                  </div>
                }
              >
                {strings.SetPropertiesLabel}
              </MessageBar>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private onErrorMessDismiss(): void {
    this.setState({ erroMessage: null });
  }
  private onSuccessMessDismiss(): void {
    this.setState({ successfulMessage: null, fileLink: null });
  }
  public getBody(): JSX.Element {
    return (
      <div className={styles.plannerReports}>
        <div className={styles.container}>
          <div>
            {this.state.loading ? <ProgressIndicator /> : null}
            {!!this.state.erroMessage ? <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={false}
              onDismiss={this.onErrorMessDismiss}
              dismissButtonAriaLabel="Close"
            >
              {this.state.erroMessage}
            </MessageBar> : null}
            {!!this.state.successfulMessage ? <MessageBar
              messageBarType={MessageBarType.success}
              isMultiline={false}
              onDismiss={this.onSuccessMessDismiss}
              actions={<Link href={this.state.fileLink} target="_blank">
              {strings.FileLabel}
            </Link>}
            >
              {this.state.successfulMessage}
            </MessageBar> : null}
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <DefaultButton onClick={this.createAndSaveDocx}
                disabled={this.state.loading || !this.state.selectedSegment}
                text={strings.CreateReportButton} />
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
  private async createAndSaveDocx() {
    this.setState({ loading: true });
    let reportRows: MicrosoftGraph.PlannerTask[];
    if (this.state.selectedTasks && this.state.selectedTasks.length > 0) {
      reportRows = this.rawTasks.filter(rt => {
        return this.state.selectedTasks.some(v => {
          return rt.id === v;
        });
      });
    } else if(this.state.onlyCompletedTaskChecked){
      reportRows = this.rawTasks.filter(v => v.percentComplete == 100);
    } else {
      reportRows = this.rawTasks;
    }

    const cellMargin = { top: 50, right: 100, bottom: 50, left: 100 };
    const tableRows: Array<TableRow> = [];
    tableRows.push(new TableRow({
      children: [
        new TableCell({
          margins: cellMargin,
          children: [new Paragraph({
            children: [
              new TextRun({
                text: "Задача",
                size: 28,
                bold: true
              })
            ]
          }
          )],
        }),
        new TableCell({
          margins: cellMargin,
          children: [new Paragraph({
            children: [
              new TextRun({
                text: "Дата завершення",
                size: 28,
                bold: true
              })
            ]
          }
          )],
        }),
        new TableCell({
          margins: cellMargin,
          children: [new Paragraph({
            children: [
              new TextRun({
                text: "Виконання",
                size: 28,
                bold: true
              })
            ]
          }
          )],
        })
      ]
    }));

    for (let task of reportRows) {
      let taskProgress;
      if (task.percentComplete === 0) {
        taskProgress = "Не розпочато";
      } else if (task.percentComplete > 0 && task.percentComplete < 100) {
        taskProgress = "На виконанні";
      } else {
        taskProgress = "Виконано";
      }
      tableRows.push(new TableRow({
        children: [
          new TableCell({
            margins: cellMargin,
            children: [new Paragraph(task.title)],
          }),
          new TableCell({
            margins: cellMargin,
            children: [new Paragraph(task.dueDateTime ? new Date(task.dueDateTime).toLocaleString() : "Не встановлено")],
          }),
          new TableCell({
            margins: cellMargin,
            children: [new Paragraph(`${taskProgress}`)],
          })
        ]
      }));
    }

    const table = new Table({
      rows: tableRows,
      width: {
        size: 100,
        type: docx.WidthType.PERCENTAGE,
      }
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
          spacing: {
            before: 200,
            after: 200
          },
          children: [
            new TextRun({
              text: this.state.selectedSegment.text,
              size: 28,
              bold: true
            })
          ],
          alignment: docx.AlignmentType.CENTER,
          pageBreakBefore: true
        }),
        table
      ]
    });
    try {
      let buf = await Packer.toBuffer(doc);
      let resp: any = await this.props.saveFile(buf, buf.byteLength,
        `${this.state.selectedSegment.text}_${date.getDate()}_${date.getMonth()}_${date.getFullYear()}.docx`);
      this.setState({ loading: false, successfulMessage: strings.FileCreatedMessage, fileLink: resp.data.LinkingUri });
    } catch (error) {
      this.setState({ erroMessage: strings.FileNotCreatedMessage });
      this.setState({ loading: false });
    }

  }
}
