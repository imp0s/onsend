import {
  base64ToUint8Array,
  removeMetadataAndComments,
  uint8ArrayToBase64
} from './docCleanup';

async function getAttachmentContent(id: string): Promise<Office.AttachmentContent> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(id, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

async function removeAndReplaceAttachment(attachment: Office.AttachmentDetailsCompose) {
  const content = await getAttachmentContent(attachment.id);
  if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
    throw new Error('Unsupported attachment format');
  }

  const cleaned = await removeMetadataAndComments(base64ToUint8Array(content.content));
  const updatedBase64 = uint8ArrayToBase64(cleaned);

  await new Promise<void>((resolve, reject) => {
    Office.context.mailbox.item.removeAttachmentAsync(attachment.id, (removeResult) => {
      if (removeResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(removeResult.error);
      }
    });
  });

  await new Promise<void>((resolve, reject) => {
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      updatedBase64,
      attachment.name,
      { isInline: attachment.isInline },
      (addResult) => {
        if (addResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(addResult.error);
        }
      }
    );
  });
}

function openConfirmationDialog(message: string): Promise<boolean> {
  return new Promise((resolve, reject) => {
    const url = `${window.location.origin}/dialog.html#${encodeURIComponent(message)}`;
    Office.context.ui.displayDialogAsync(url, { height: 40, width: 30 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
        reject(result.error);
        return;
      }

      const dialog = result.value;
      const handler = (arg: Office.DialogMessageReceivedEventArgs) => {
        dialog.removeEventHandler(Office.EventType.DialogMessageReceived, handler);
        dialog.close();
        resolve(arg.message === 'yes');
      };

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, handler);
    });
  });
}

export async function promptForCleanup(names: string[]): Promise<boolean> {
  const message = `Clean metadata and comments from: ${names.join(', ')}?`;
  if (typeof Office !== 'undefined' && Office.context?.ui?.displayDialogAsync) {
    try {
      return await openConfirmationDialog(message);
    } catch (error) {
      console.error('Dialog failed; defaulting to proceed', error);
      return true;
    }
  }

  return typeof window !== 'undefined' ? window.confirm(message) : true;
}

export async function onMessageSend(event: Office.AddinCommands.Event) {
  try {
    const attachments = Office.context.mailbox.item.attachments || [];
    const targets = attachments.filter((att) => att.name.toLowerCase().endsWith('.docx'));

    if (targets.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    const shouldClean = await promptForCleanup(targets.map((a) => a.name));
    if (!shouldClean) {
      event.completed({ allowEvent: true });
      return;
    }

    for (const attachment of targets) {
      await removeAndReplaceAttachment(attachment);
    }

    event.completed({ allowEvent: true });
  } catch (error) {
    console.error('Cleanup failed, sending without modification', error);
    event.completed({ allowEvent: true });
  }
}

Office.actions.associate('onMessageSend', onMessageSend);
