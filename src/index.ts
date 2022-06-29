import fs from 'fs';
import { AttachmentData } from '@sendgrid/helpers/classes/attachment';
import { MailDataRequired } from '@sendgrid/helpers/classes/mail';
import sgMail, { MailService } from '@sendgrid/mail';
import {
  DatasourceConfiguration,
  DatasourceMetadataDto,
  EMAIL_INTEGRATION_SENDER_ADDRESS_DEFAULT,
  EMAIL_INTEGRATION_SENDER_NAME_DEFAULT,
  EmailActionConfiguration,
  EmailActionFieldNames,
  EmailActionFieldsMap,
  EmailDatasourceConfiguration,
  ExecutionOutput,
  FormItem,
  IntegrationError,
  isReadableFile,
  isReadableFileConstructor,
  NotFoundError,
  RawRequest
} from '@superblocksteam/shared';
import { BasePlugin, PluginExecutionProps, RequestFile, RequestFiles } from '@superblocksteam/shared-backend';
import { isEmpty } from 'lodash';

export default class EmailPlugin extends BasePlugin {
  async execute({
    datasourceConfiguration,
    actionConfiguration,
    files
  }: PluginExecutionProps<EmailDatasourceConfiguration>): Promise<ExecutionOutput> {
    for (const field of Object.values<FormItem>(EmailActionFieldsMap)) {
      if (field.rules?.[0]?.required) {
        if (!actionConfiguration[field.name]) {
          throw new IntegrationError(`${field.label} not specified`);
        }
      }
    }

    const msg = this.formEmailJson(actionConfiguration, files);

    try {
      const client = this.createClient(datasourceConfiguration);

      await client.send({
        ...msg,
        from: {
          email: datasourceConfiguration.authentication?.custom?.senderEmail?.value ?? EMAIL_INTEGRATION_SENDER_ADDRESS_DEFAULT,
          name: datasourceConfiguration.authentication?.custom?.senderName?.value ?? EMAIL_INTEGRATION_SENDER_NAME_DEFAULT
        }
      });
      const ret = new ExecutionOutput();
      ret.output = msg;
      return ret;
    } catch (err) {
      throw new IntegrationError(`Failed to send email using SendGrid.\n\nError:\n${err}`);
    }
  }

  createClient(datasourceConfiguration: EmailDatasourceConfiguration): MailService {
    const key = datasourceConfiguration.authentication?.custom?.apiKey?.value ?? '';
    if (isEmpty(key)) {
      throw new NotFoundError('No API key found for Email integration');
    }

    sgMail.setApiKey(key);
    return sgMail;
  }

  dynamicProperties(): string[] {
    return Object.values(EmailActionFieldNames);
  }

  formEmailJson(actionConfiguration: EmailActionConfiguration, files: RequestFiles = undefined): MailDataRequired {
    let attachments: AttachmentData[];
    if (actionConfiguration.emailAttachments) {
      if (typeof actionConfiguration.emailAttachments === 'string') {
        try {
          actionConfiguration.emailAttachments = JSON.parse(actionConfiguration.emailAttachments);
        } catch (e) {
          throw new IntegrationError(`Can't parse the file objects. They must be an array of JSON objects.`);
        }
      }

      if (!Array.isArray(actionConfiguration.emailAttachments)) {
        throw new IntegrationError(`Attachments must be provided in the form of an array of JSON objects.`);
      }

      attachments = actionConfiguration.emailAttachments.map((file: unknown) => {
        // Check if the object being passed is a Superblocks file
        // object or has properties that allow it to be read as one
        if (!isReadableFile(file)) {
          if (isReadableFileConstructor(file)) {
            // Sendgrid requires the attached file content to be base64 encoded
            return { filename: file.name, content: Buffer.from(file.contents).toString('base64'), type: file.type } as AttachmentData;
          }

          throw new IntegrationError(
            'Cannot read attachments. Attachments can either be Superblocks files or { name: string; contents: string, type: string }.'
          );
        }

        const match = (files as Array<RequestFile>).find((f) => f.filename === file.$superblocksId);
        if (!match) {
          throw new IntegrationError(`Could not locate contents for attachment file ${file.name}`);
        }
        return { filename: file.name, content: fs.readFileSync(match.path).toString('base64'), type: file.type } as AttachmentData;
      });
    }

    return {
      from: actionConfiguration.emailFrom,
      to: this.parseEmailAddresses(actionConfiguration.emailTo),
      cc: this.parseEmailAddresses(actionConfiguration.emailCc),
      bcc: this.parseEmailAddresses(actionConfiguration.emailBcc),
      subject: actionConfiguration.emailSubject,
      html: actionConfiguration.emailBody,
      attachments: attachments
    };
  }

  getRequest(
    actionConfiguration: EmailActionConfiguration,
    datasourceConfiguration: DatasourceConfiguration,
    files: RequestFiles
  ): RawRequest {
    return JSON.stringify(this.formEmailJson(actionConfiguration, files), null, 2);
  }

  parseEmailAddresses(emailsString: string): string[] {
    if (isEmpty(emailsString)) {
      return [];
    }
    // Trim any whitespace and remove any empty strings from the split
    return emailsString
      .split(',')
      .map((item) => item.trim())
      .filter((item) => !isEmpty(item));
  }

  async metadata(datasourceConfiguration: EmailDatasourceConfiguration): Promise<DatasourceMetadataDto> {
    return {};
  }

  async test(datasourceConfiguration: EmailDatasourceConfiguration): Promise<void> {
    return;
  }
}
