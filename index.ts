import "dotenv/config";
import { env } from "process";
import * as fs from "fs";
import { PSTAttachment, PSTFile, PSTFolder, PSTMessage } from "pst-extractor";
//@ts-ignore
import emlFormat from "eml-format";
import util from "util";

import { sleep } from "./common";
import { uploadContent } from "./s3";

const INSERT_SIZE = 10;
const SLEEP = 1000 * 60;

// Default buffer size each time read from stream
const BUFFER_SIZE = 8176;

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
  subject: string;
  text: string;
  html: string;
  attachments?: Attachment[];
};

//@TODO: Update it to your PST file path
const pstFolder = "/Users/tranvinhphuc/scripts/streamPSTFile/input/backup.pst";
const saveToFS = true;

const verbose = true;
const displaySender = true;
const displayBody = true;
let depth = -1;
let col = 0;

// console log highlight with https://en.wikipedia.org/wiki/ANSI_escape_code
const ANSI_RED = 31;
const ANSI_YELLOW = 93;
const highlight = (str: string, code: number = ANSI_RED) =>
  "\u001b[" + code + "m" + str + "\u001b[0m";

// eml
const buildEml = util.promisify(emlFormat.build);

/**
 * Returns a string with visual indication of depth in tree.
 * @param {number} depth
 * @returns {string}
 */
const getDepth = (depth: number): string => {
  let sdepth = "";
  if (col > 0) {
    col = 0;
    sdepth += "\n";
  }
  for (let x = 0; x < depth - 1; x++) {
    sdepth += " | ";
  }
  sdepth += " |- ";
  return sdepth;
};

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
    const buffer = Buffer.alloc(attachment.size);
    let bytesRead = 0;
    do {
      bytesRead = attachmentStream.read(buffer);
    } while (bytesRead == BUFFER_SIZE);

    attachments.push({
      name: attachment.filename,
      data: buffer,
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
  if (env.ENV === "dev") {
    const path = `${year}-${month}-${day}-${uid}.eml`;
    fs.writeFileSync(
      `/Users/tranvinhphuc/scripts/streamPSTFile/output/${path}`,
      eml
    );
  } else {
    const path = `${env.COMPANY_ID}/${env.COMPANY_EMAIL}/${year}-${month}-${day}/${uid}.eml`;
    await uploadContent({
      remotePath: path,
      content: eml,
      metaData: {
        group_id: env.GROUP_ID,
        company_id: env.COMPANY_ID,
      },
    });
  }
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
  if (verbose && displaySender && email.messageClass === "IPM.Note") {
    console.log(getDepth(depth) + " sender: " + sender);
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
 * Print a dot representing a message.
 */
const printDot = (): void => {
  process.stdout.write(".");
  if (col++ > 100) {
    console.log("");
    col = 0;
  }
};

/**
 * Walk the folder tree recursively and process emails.
 * @param {PSTFolder} folder
 */
const processFolder = async (folder: PSTFolder): Promise<void> => {
  depth++;
  let attempt = 0;

  // the root folder doesn't have a display name
  if (depth > 0) {
    console.log(getDepth(depth) + folder.displayName);
  }

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
      attempt++;
      if (verbose) {
        console.log(
          getDepth(depth) +
            "Email: " +
            email.descriptorNodeId +
            " - " +
            email.subject
        );
      } else {
        printDot();
      }

      // sender
      const sender = getSender(email);

      // recipients
      const recipients = getRecipients(email);

      // display body?
      if (verbose && displayBody) {
        console.log(highlight("email.body", ANSI_YELLOW), email.body);
        console.log(highlight("email.bodyRTF", ANSI_YELLOW), email.bodyRTF);
        console.log(highlight("email.bodyHTML", ANSI_YELLOW), email.bodyHTML);
      }

      // save content to fs?
      if (saveToFS) {
        // create date string in format YYYY-MM-DD
        let d = email.clientSubmitTime;
        if (!d && email.creationTime) {
          d = email.creationTime;
        }

        await doSaveToFS(email, sender, recipients);
        if (attempt % (INSERT_SIZE + 1) === INSERT_SIZE) {
          await sleep(SLEEP);
        }
      }
      email = folder.getNextChild();
    }
    depth--;
  }
  depth--;
};
// load file into memory buffer, then open as PSTFile
const pstFile = new PSTFile(pstFolder);
console.log(pstFile.getMessageStore().displayName);
processFolder(pstFile.getRootFolder()).catch(console.error);
