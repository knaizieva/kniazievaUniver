import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

export interface IPlannerReportsProps {
  description: string;
  group:string;
  plan:string;
  library:string;
  getTasks: (planId:string) => Promise<MicrosoftGraph.PlannerTask[]>;
  getTaskDetails: (taskId:string) => Promise<MicrosoftGraph.PlannerTaskDetails>;
  getBuckets: (planId:string) => Promise<MicrosoftGraph.PlannerBucket[]>;
  getBucketTasks: (bucketId: string | number) => Promise<MicrosoftGraph.PlannerTask[]>;
}
