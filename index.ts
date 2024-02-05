require("dotenv").config();
import { PSTAttachment, PSTFile, PSTFolder, PSTMessage } from "pst-extractor";
//@ts-ignore
import emlFormat from "eml-format";
import util from "util";
import os from "os";
import BPromise from "bluebird";
import path from "path";
import { getDuplicatedEmails, getEmlHandler, logger, sleep } from "./utils/common";
import { partition } from "lodash";

const INSERT_SIZE = 200;
const S3_BATCH_SIZE = 10;
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
const FILE_TO_PROCESS = "BIZZI.1.DONE.pst";
const WORK_DIR = os.tmpdir();
const LOG_FILE = path.join(WORK_DIR, path.parse(FILE_TO_PROCESS).name + ".txt");
const pstFolder = "/home/bizzivietnam/masan";

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
  // const newEml = msg.transportMessageHeaders.concat(eml);  
  const filePath = path.join(`${year}-${month}-${day}`, `${uid}_${Date.now()}.eml`);

  const destination = await emailHandler.handle(eml, filePath);
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
  let emails:  PSTMessage[] = [];
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
    let email: PSTMessage = folder.getNextChild();
    while (email != null) {
      processEmailCount++;
      emails.push(email)

      if (processEmailCount % (INSERT_SIZE + 1) !== INSERT_SIZE) {
        email = folder.getNextChild();
        continue;
      }
      console.time("DUPLICATED");
      const duplicatedEmails = await getDuplicatedEmails(
        getSender(email),
        emails.map(email => email.subject),
      );
      console.timeEnd("DUPLICATED");
      
      const duplicatedEmailMappers = new Map(
        duplicatedEmails.map(duplicatedEmail => {
          const extractEmlPattern = /\/([A-Za-z0-9_=\-]+\.eml)/i;
          const emlFileName = duplicatedEmail.object.key.match(extractEmlPattern)?.[1];
          const id =  path.parse(emlFileName).name.split("_")[0];
          return [id, true]
        }
      ));

      const [nonDuplicatedEmails, duplicatedEmailPartitions] = partition(emails, 
        (email) => !duplicatedEmailMappers.has(email.descriptorNodeId.toString())
      );

      await BPromise.map(
        nonDuplicatedEmails,
        async email => {
          // sender
          const sender = getSender(email);
          // recipients
          const recipients = getRecipients(email);
          await doSaveToFS(email, sender, recipients);
          const message = `Processed: ${processEmailCount}/${totalEmailCount} - Descriptor Node ID: ${email.descriptorNodeId}`
          await logger(message, LOG_FILE);
        },
        {
          concurrency: S3_BATCH_SIZE
        }
      );

      await BPromise.map(
        duplicatedEmailPartitions,
        async email => {
          const message = `Duplicated email with descriptor Node ID: ${email.descriptorNodeId}`
          await logger(message, LOG_FILE);
        }
      )

      // clean-up
      emails = [];
      await sleep(SLEEP);
    }
  }
};
// load file into memory buffer, then open as PSTFile
const pstFile = new PSTFile(
  path.join(pstFolder, FILE_TO_PROCESS)
);
console.log(pstFile.getMessageStore().displayName);
processFolder(pstFile.getRootFolder()).catch(console.error);
