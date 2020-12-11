import * as chalk from 'chalk';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  contentType: string;
  startTime?: string;
  endTime?: string;
}

interface ActivityfeedSubscription {
  contentType: string;
  status: string;
  webhook: string;
}

interface AuditContentList {
  contentType: string;
  contentId: string;
  contentUri: string;
  contentCreated: string;
  contentExpiration: string;
}

interface AuditlogReport {
  CreationTime: string;
  Id: string;
  Workload: string;
  Operation: string;
  ClientIP: string;
  User: string;
}

enum AuditContentTypes {
  AzureActiveDirectory = "Audit.AzureActiveDirectory",
  Exchange = "Audit.Exchange",
  SharePoint = "Audit.SharePoint",
  General = "Audit.General ",
  DLP = "DLP.All"
}
class TenantAuditlogReportCommand extends Command {
  private serviceUrl: string = 'https://manage.office.com/api/v1.0';
  private tenantId: string | undefined;
  public get name(): string {
    return `${commands.TENANT_AUDITLOG_REPORT}`;
  }

  public get description(): string {
    return 'Gets audit logs from the Office 365 Management API';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.startTime = args.options.startTime;
    telemetryProps.endTime = args.options.endTime;
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['Operation', 'Id', 'UserId', 'Workload', 'ClientIP'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Start retrieving Audit Log Report`);
    }

    this.tenantId = Utils.getTenantIdFromAccessToken(auth.service.accessTokens[auth.defaultResource].value);
    this.startContentSubscriptionifNotActive(args, logger)
      .then((): Promise<AuditContentList[]> => this.getAuditContentList(args, logger))
      .then((AuditContentLists: AuditContentList[]): Promise<any> => this.getBatchedAuditLogReports(logger,AuditContentLists))
      .then((res: any): void => {

        // logger.log(res);

        // logger.log(`Total Audit Log Count : ${res.length}`)

        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));

  }

  private startContentSubscriptionifNotActive(args: CommandArgs, logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Check whether subscription is active.`);
    }
    let SubscriptionListEndpoint: string = 'activity/feed/subscriptions/list';

    const requestOptions: any = {
      url: `${this.serviceUrl}/${this.tenantId}/${SubscriptionListEndpoint}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<ActivityfeedSubscription[]>(requestOptions)
      .then((subscriptionLists: ActivityfeedSubscription[]): boolean => {
        return subscriptionLists.some(subscriptionList => subscriptionList
          .contentType === (<any>AuditContentTypes)[args.options.contentType]);
      })
      .then((hasActiveSubscription: boolean): Promise<void> => {
        if (!hasActiveSubscription) {
          if (this.verbose) {
            logger.logToStderr(`Starting subscription since subscription is not active for the content type`);
          }
          let startSubscriptionEndPoint: string = `activity/feed/subscriptions/start?contentType=${(<any>AuditContentTypes)[args.options.contentType]}&PublisherIdentifier=${this.tenantId}`;
          const requestOptions: any = {
            url: `${this.serviceUrl}/${this.tenantId}/${startSubscriptionEndPoint}`,
            headers: {
              accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
          };

          return request.post(requestOptions);
        }

        return Promise.resolve();
      });
  }

  private getAuditContentList(args: CommandArgs, logger: Logger): Promise<AuditContentList[]> {
    if (this.verbose) {
      logger.logToStderr(`Start listing Audit Content URL`);
    }

    let SubscriptionListEndpoint: string = (typeof args.options.startTime !== 'undefined' && typeof args.options.endTime !== 'undefined') ?
      `/activity/feed/subscriptions/content?contentType=${(<any>AuditContentTypes)[args.options.contentType]}&PublisherIdentifier=${this.tenantId}&starttime=${args.options.startTime}&endTime=${args.options.endTime}` :
      `/activity/feed/subscriptions/content?contentType=${(<any>AuditContentTypes)[args.options.contentType]}&PublisherIdentifier=${this.tenantId}`;
    const requestOptions: any = {
      url: `${this.serviceUrl}/${this.tenantId}/${SubscriptionListEndpoint}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<AuditContentList[]>(requestOptions)
  }

  private getBatchedAuditLogReports(logger: Logger, AuditContentLists: AuditContentList[]) : Promise<any> {
    if (this.verbose) {
      logger.logToStderr(`Start generating Audit Reports in batchwise manner`);
    }

    //Arjun - New Approach

    const batchSize: number = 10;
    //Temporary - Will be removed in production
    let tempMaximumCount: number = 20;

    let BatchedAuditContentList: any = [];

    for (let i = 0; i < (AuditContentLists.length < tempMaximumCount ? AuditContentLists.length : tempMaximumCount); i += batchSize) {
      const AuditContent : AuditContentList[] = AuditContentLists.slice(i, i + batchSize < AuditContentLists.length ? i + batchSize : AuditContentLists.length).map((AuditContentList) => {
        //logger.log(`Inner Loop : ${i}`);        
        return AuditContentList;
      })
      BatchedAuditContentList.push(AuditContent);
    }

    logger.log(`Total Batches : ${BatchedAuditContentList.length}`)
    logger.log(`Batched Result is : ${BatchedAuditContentList[0][0]}`)

    return this.getAuditLogReportsforCompleteBatch(logger, BatchedAuditContentList)
      .then((CompleteAuditReportsUnflattened: any): Promise<any> => {

        logger.log(`Result is : ${CompleteAuditReportsUnflattened}`)
        const CompleteAuditlogReports = CompleteAuditReportsUnflattened.flat(2);
        return Promise.resolve(CompleteAuditlogReports);
      });


    //End of New Approach
    
    // //Batching Approach - Given Batch size is 10
    // const batchSize: number = 10;
    // let runningBatch: number = 0;
    // var CompleteAuditReportsUnflattened : any = [];
    // //Temporary - Will be removed in production
    // let tempMaximumCount: number = 20;

    // for (let i = 0; i < (AuditContentLists.length<tempMaximumCount?AuditContentLists.length:tempMaximumCount); i += batchSize) {
    //   if (this.verbose) {
    //     runningBatch++;
    //     logger.logToStderr(`Generating Audit Reports for Batch : ${runningBatch}`);
    //   }

    //   const requests = AuditContentLists.slice(i, i + batchSize<AuditContentLists.length?i+batchSize:AuditContentLists.length).map((AuditContentList) => {
    //     //logger.log(`Inner Loop : ${i}`);        
    //     return this.getAuditLogReportforSingleContentURL(logger, AuditContentList.contentUri);
    //   })
      
    //   // //Async Approach
    //   // let batchedAuditReport : any  = [];
    //   // batchedAuditReport = await Promise.all(requests);
    //   // CompleteAuditReportsUnflattened.push(batchedAuditReport);

    //   //NOT Async Approach
    //   Promise.all(requests)
    //   .then((batchedAuditReport : any): void => {
    //     logger.log(`Inside THEN Method. Let us if this works`);
    //     logger.log(`${batchedAuditReport[0][0]}`);
    //     CompleteAuditReportsUnflattened.push(batchedAuditReport);
    //   });
    // }
    
    // //logger.log(`Just before final push. Let us see how this behaves`);
    // const CompleteAuditlogReports = CompleteAuditReportsUnflattened.flat(2);
    // return Promise.resolve(CompleteAuditlogReports);
  }

  private getAuditLogReportforSingleContentURL(logger: Logger, auditURL: string): Promise<AuditlogReport[]> {
    return new Promise<AuditlogReport[]>((resolve: (AuditLogs: AuditlogReport[]) => void, reject: (error: any) => void): void => {
      const requestOptions: any = {
        url: auditURL,
        headers: {
          accept: 'application/json;'
        },
        responseType: 'json'
      };

      logger.log(`Inside the Generate Audit Section`);

      request
        .get<AuditlogReport[]>(requestOptions)
        .then((AuditLogs: AuditlogReport[]): void => {
          logger.log(`Audit Log Resolved`)
          resolve(AuditLogs);
        }, (err: any): void => {
          logger.log(`Audit Log Rejected`)
          reject(err);
        });
    });
  }

  //Arjun's New Method

  // private getBatchedAuditContent(logger: Logger,AuditContentLists : AuditContentList[],  batchSize: number = 10): AuditContentList[][] {
  //   let runningBatch: number = 0;
  //   //Temporary - Will be removed in production
  //   let tempMaximumCount: number = 20;

  //   const CompleteRequest: any = [];

  //   for (let i = 0; i < (AuditContentLists.length < tempMaximumCount ? AuditContentLists.length : tempMaximumCount); i += batchSize) {
  //     if (this.verbose) {
  //       runningBatch++;
  //       logger.logToStderr(`Generating Audit Reports for Batch : ${runningBatch}`);
  //     }

  //     const AuditContent : AuditContentList[] = AuditContentLists.slice(i, i + batchSize < AuditContentLists.length ? i + batchSize : AuditContentLists.length).map((AuditContentList) => {
  //       //logger.log(`Inner Loop : ${i}`);        
  //       return AuditContentList;
  //     })

  //     CompleteRequest.push(AuditContent);
  //   }

  //   return CompleteRequest;
  // }

  private getAuditLogReportsforCompleteBatch(logger: Logger, CompleteAuditContentList: AuditContentList[][]): Promise<any> {
    var results : any = [];
    var index = 0;

    const GenerateBatchedAuditReports = (): Promise<any> => {
      if(this.verbose){
        logger.log(`Get Audit Report for Batch : ${index}`)
      }
      if (index < CompleteAuditContentList.length) {
        return this.getAuditContentsforSingleBatch(logger, CompleteAuditContentList[index++])
        .then((data: any): void => {
          results.push(data);
          GenerateBatchedAuditReports();
        });
      } else {
        return Promise.resolve(results);
      }
    };

    GenerateBatchedAuditReports();

    return Promise.resolve();
  }

  private getAuditContentsforSingleBatch(logger : Logger, SingleBatchAuditContentList: AuditContentList[]) : Promise<any> {
    return Promise.all(SingleBatchAuditContentList.map(AuditContentInfo => {
      return this.getAuditLogReportforSingleContentURL(logger,AuditContentInfo.contentUri)
    }));
  }

  //End of Arjun's MEthod

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-c, --contentType <contentType>',
        description: 'Audit content type of logs to be retrieved, should be one of the following: AzureActiveDirectory, Exchange, SharePoint, General, DLP'
      },
      {
        option: '-s, --startTime [startTime]',
        description: 'Start time of logs to be retrieved. Start time and end time must both be specified (or both omitted) and must be less than or equal to 24 hours apart.'
      },
      {
        option: '-e, --endTime [endTime]',
        description: 'End time of logs to be retrieved. Start time and end time must both be specified (or both omitted) and must be less than or equal to 24 hours apart.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {

    if ((<any>AuditContentTypes)[args.options.contentType] === undefined) {
      return `${args.options.contentType} is not a valid value for the contentType option. Allowed values are ${Object.keys(AuditContentTypes).join(' | ')}`;
    }

    return true;
  }

}

module.exports = new TenantAuditlogReportCommand();