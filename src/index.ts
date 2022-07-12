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
  ExecutionContext,
  ExecutionOutput,
  FormItem,
  IntegrationError,
  isReadableFile,
  isReadableFileConstructor,
  NotFoundError,
  RawRequest
} from '@superblocksteam/shared';
import { BasePlugin, PluginExecutionProps, RequestFile, RequestFiles, getEncodedFile } from '@superblocksteam/shared-backend';
import { isEmpty } from 'lodash';

export default class EmailPlugin extends BasePlugin {
  async execute({
    context,
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

    const msg = await this.formEmailJson(context, actionConfiguration, files);

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

  async formEmailJson(
    context: ExecutionContext,
    actionConfiguration: EmailActionConfiguration,
    files: RequestFiles = undefined
  ): Promise<MailDataRequired> {
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

      attachments = await Promise.all(
        actionConfiguration.emailAttachments.map(async (file: unknown) => {
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

          const match = (files as Array<RequestFile>).find((f) => f.filename.startsWith(`${file.$superblocksId}_`));
          if (!match) {
            throw new IntegrationError(`Could not locate contents for attachment file ${file.name}`);
          }
          try {
            return {
              filename: file.name,
              content: await getEncodedFile(context, match.path, 'base64'),
              type: file.type
            } as AttachmentData;
          } catch (_) {
            throw new IntegrationError(`Could not retrieve file ${file.name} from controller.`);
          }
        })
      );
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
    // Can't call formEmailJson anymore because it's async and because
    // we probs shouln't be grabbing the attachment data anyways.
    // We'll need to add some attachment metadata back in the future.
    return JSON.stringify(
      {
        from: actionConfiguration.emailFrom,
        to: this.parseEmailAddresses(actionConfiguration.emailTo),
        cc: this.parseEmailAddresses(actionConfiguration.emailCc),
        bcc: this.parseEmailAddresses(actionConfiguration.emailBcc),
        subject: actionConfiguration.emailSubject,
        html: actionConfiguration.emailBody
      },
      null,
      2
    );
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
