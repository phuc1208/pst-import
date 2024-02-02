require("dotenv").config();
import { PSTAttachment, PSTFile, PSTFolder, PSTMessage } from "pst-extractor";
//@ts-ignore
import emlFormat from "eml-format";
import util from "util";
import os from "os";

import { getEmlHandler, logger, sleep } from "./common";
import path from "path";

const INSERT_SIZE = 100;
const SLEEP = 1000 * 60 * 5;

type Attachment = {
  name: string;
  contentType: string;
  data: string | Buffer;
};

type Email = {
  from: string;
  to: {
    name: string;
    email: string;
  };
  headers: Record<string, any>,
  subject: string;
  text: string;
  html: string;
  attachments?: Attachment[];
};


//@TODO: Update it to your PST file path
const FILE_TO_PROCESS = "backup.pst";
const WORK_DIR = os.tmpdir();
const LOG_FILE = path.join(WORK_DIR, path.parse(FILE_TO_PROCESS).name + ".txt");
const pstFolder = "/Users/tranvinhphuc/scripts/streamPSTFile/input";


let depth = -1;

const emailHandler = getEmlHandler();

// eml
const buildEml = util.promisify(emlFormat.build);

/**
 * Save items to filesystem.
 * @param {PSTMessage} msg
 * @param {string} emailFolder
 * @param {string} sender
 * @param {string} recipients
 */
const doSaveToFS = async (
  msg: PSTMessage,
  sender: string,
  recipients: string
): Promise<void> => {
  const email: Email = {
    from: "",
    to: {
      name: "",
      email: "",
    },
    headers: {},
    subject: "",
    text: "",
    html: "",
    attachments: [],
  };
  email.from = sender;
  email.to = {
    name: recipients,
    email: msg.displayTo,
  };
  email.text = msg.body;
  email.html = msg.bodyHTML;
  email.subject = msg.subject;

  // walk list of attachments and save to fs
  const attachments: Attachment[] = [];
  for (let i = 0; i < msg.numberOfAttachments; i++) {
    const attachment: PSTAttachment = msg.getAttachment(i);
    // Log.debug1(JSON.stringify(activity, null, 2));
    if (!attachment.filename) {
      continue;
    }

    const attachmentStream = attachment.fileInputStream;
    if (!attachmentStream) {
      continue;
    }
    const bufferSize = 8176;
    const buffer = Buffer.alloc(bufferSize);
    let mergeBuffer = Buffer.alloc(0);
    let bytesRead = 0;
    do {
      bytesRead = attachmentStream.read(buffer);
      mergeBuffer = Buffer.concat([mergeBuffer, buffer.slice(0, bytesRead)]);
    } while (bytesRead == bufferSize);

    attachments.push({
      name: attachment.longFilename,
      data: mergeBuffer,
      contentType: attachment.mimeTag,
    });
  }

  const deliveryDate = msg.messageDeliveryTime || new Date();
  const year = deliveryDate.getFullYear();
  const month = `${deliveryDate.getMonth() + 1}`.padStart(2, "0");
  const day = `${deliveryDate.getDate()}`.padStart(2, "0");
  const uid = msg.descriptorNodeId;
  email.attachments = attachments;

  const eml = await buildEml(email);
  const newEml = msg.transportMessageHeaders.concat(eml);  
  const filePath = path.join(`${year}-${month}-${day}`, `${uid}.eml`);

  const destination = await emailHandler.handle(newEml, filePath);
  console.log(`Save email to ${destination}`);
};

/**
 * Get the sender and display.
 * @param {PSTMessage} email
 * @returns {string}
 */
const getSender = (email: PSTMessage): string => {
  let sender = email.senderName;
  if (sender !== email.senderEmailAddress) {
    sender += " (" + email.senderEmailAddress + ")";
  }
  return sender;
};

/**
 * Get the recipients and display.
 * @param {PSTMessage} email
 * @returns {string}
 */
const getRecipients = (email: PSTMessage): string => {
  // could walk recipients table, but be fast and cheap
  return email.displayTo;
};

/**
 * Walk the folder tree recursively and process emails.
 * @param {PSTFolder} folder
 */
const processFolder = async (folder: PSTFolder): Promise<void> => {
  depth++;

  const totalEmailCount = folder.emailCount;
  let processEmailCount = 0;

  // go through the folders...
  if (folder.hasSubfolders) {
    const childFolders: PSTFolder[] = folder.getSubFolders();
    for (const childFolder of childFolders) {
      processFolder(childFolder);
    }
  }

  // and now the emails for this folder
  if (folder.contentCount > 0) {
    depth++;
    let email: PSTMessage = folder.getNextChild();
    while (email != null) {
      processEmailCount++;

      // sender
      const sender = getSender(email);

      // recipients
      const recipients = getRecipients(email);

      // create date string in format YYYY-MM-DD
      let d = email.clientSubmitTime;
      if (!d && email.creationTime) {
        d = email.creationTime;
      }

      await doSaveToFS(email, sender, recipients);
      const message = `Processed: ${processEmailCount}/${totalEmailCount} - Descriptor Node ID: ${email.descriptorNodeId}`;
      await logger(message, LOG_FILE);
      break;
      if (processEmailCount % (INSERT_SIZE + 1) === INSERT_SIZE) {
        await sleep(SLEEP);
      }

      email = folder.getNextChild();
    }
    depth--;
  }
  depth--;
};
// load file into memory buffer, then open as PSTFile
const pstFile = new PSTFile(
  path.join(pstFolder, FILE_TO_PROCESS)
);
console.log(pstFile.getMessageStore().displayName);
processFolder(pstFile.getRootFolder()).catch(console.error);
