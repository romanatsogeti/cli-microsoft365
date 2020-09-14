//import config from '../../../../config';
import commands from '../../commands';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandError, CommandOption, CommandValidate } from '../../../../Command';
import { ContextInfo } from '../../spo';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  sourceUrl?: string;
  targetUrl?: string;
  archiveUrl?: string;
  disableRedirection?: boolean;
  wait?: boolean;
}

class SpoSiteSwapCommand extends SpoCommand {

  public get name(): string {
    return commands.SITE_SWAP;
  }

  public get description(): string {
    return 'Swap the location of a site with another site while archiving the original site';
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--sourceUrl <sourceUrl>',
        description: 'URL of the source site'
      },
      {
        option: '--targetUrl <targetUrl>',
        description: 'URL of the target site that the source site will be swapped to'
      },
      {
        option: '--archiveUrl <archiveUrl',
        description: 'URL that the target site will be archived to'
      },
      {
        option: '--disableRedirection',
        description: 'Disables the site redirect from being created at the Source URL location'
      },
      {
        option: '--wait',
        description: 'Wait for the job to complete'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.sourceUrl) {
        return 'Required source url is missing';
      } else {
        const isValidSharePointSourceUrl : boolean | string  = SpoCommand.isValidSharePointUrl(args.options.sourceUrl);

        if (isValidSharePointSourceUrl !== true) {
          return isValidSharePointSourceUrl;
        }
      }

      if (!args.options.targetUrl) {
        return 'Required target url is missing';
      } else {
        const isValidSpoTargetUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.targetUrl);
        if (isValidSpoTargetUrl !== true) {
          return isValidSpoTargetUrl;
        }
      }

      if (!args.options.archiveUrl) {
        return 'Required archive url is missing';
      } else {
        const isValidSpoArchiveUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.targetUrl);
        if (isValidSpoArchiveUrl !== true) {
          return isValidSpoArchiveUrl;
        }
      }

      return true;
    };
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string;

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;

        return this.getRequestDigest(spoAdminUrl);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void =>  {
        if (err instanceof CommandError) {
          err = (err as CommandError).message;
        }

        this.handleRejectedPromise(err, cmd, cb)
      });
    

    let _rawQuery: string = ` 
    '<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="SharePoint Online PowerShell (16.0.20414.0)" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
      <Actions>
        <ObjectPath Id="4" ObjectPathId="3" />
        <ObjectPath Id="6" ObjectPathId="5" />
        <Query Id="7" ObjectPathId="5">
          <Query SelectAllProperties="true">
            <Properties />
          </Query>
        </Query>
      </Actions>
      <ObjectPaths>
        <Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" />
        <Method Id="5" ParentId="3" Name="SwapSiteWithSmartGestureOptionForce">
          <Parameters>
            <Parameter Type="String">https://rvrcapgemini.sharepoint.com/sites/work</Parameter>
            <Parameter Type="String">https://rvrcapgemini.sharepoint.com/sites/wlive</Parameter>
            <Parameter Type="String">https://rvrcapgemini.sharepoint.com/sites/wlive-archive</Parameter>
            <Parameter Type="Boolean">true</Parameter>
            <Parameter Type="Boolean">false</Parameter>
          </Parameters>
        </Method>
      </ObjectPaths>
    </Request>'`;
    
    cmd.log(_rawQuery);
  } 

  public commandHelp(args: any, log: (message: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.SITE_SWAP).helpInformation());

    log(
      `  ${chalk.yellow('Important:')} to use this command you have to have permissions to access
    the tenant admin site.

    Remarks:
      The source and target sites can't be connected to an Office 365 group. They also can't be hub sites or associated with a hub. 
      If a site is a hub site, unregister it as a hub site, swap the root site, and then register the site as a hub site. 
      If a site is associated with a hub, disassociate the site, swap the root site, and then reassociate the site.

    `);
  }
}

module.exports = new SpoSiteSwapCommand();