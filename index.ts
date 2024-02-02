import "dotenv/config";
import { env } from "process";
import * as fs from 'fs'
import { PSTAttachment, PSTFile, PSTFolder, PSTMessage } from 'pst-extractor'
//@ts-ignore
import emlFormat from "eml-format";
import util from "util";
import { S3 } from '@aws-sdk/client-s3';
import type { ResponseMetadata } from "@aws-sdk/types";

const INSERT_SIZE = 10;
const SLEEP = 1000*60;
let attempt = 0;

type Attachment = {
  name: string;
  contentType: string;
  data: string | Buffer;
}
type Email = {
  from: string;
  to: {
      name: string;
      email: string;
  };
  subject: string;
  text: string;
  html: string;
  attachments?: Attachment[]
}
type ApiResponse<T = any, U = Error> = { data: T; err: U };

const pstFolder = '/Users/tranvinhphuc/scripts/streamPSTFile/input/backup.pst'
const saveToFS = true

const verbose = true
const displaySender = true
const displayBody = true
let depth = -1
let col = 0

// console log highlight with https://en.wikipedia.org/wiki/ANSI_escape_code
const ANSI_RED = 31
const ANSI_YELLOW = 93
const highlight = (str: string, code: number = ANSI_RED) => '\u001b[' + code + 'm' + str + '\u001b[0m'

// eml
const buildEml = util.promisify(emlFormat.build);
const sleep = (time: number) => new Promise(res => {
  console.info(`wait for ${time} seconds`);
  setTimeout(() => {
    console.info(`complete wait for ${time} seconds`)
    res(time);
  }, time);
});

// s3
const s3 = new S3({ maxAttempts: 3, region: env.REGION });
export const uploadContent = async ({
  remotePath: objectKey,
  content,
  metaData,
}: {
  remotePath: string;
  content: string;
  metaData?: Record<string, string>;
}): Promise<ApiResponse<ResponseMetadata>> => {
  try {
    const response = await s3.putObject({
      Bucket: env.BUCKET,
      Body: content,
      Key: objectKey,
      Metadata: metaData,
    });
    return { data: response.$metadata, err: null };
  } catch (err: any) {
    console.error("Upload file err", err);
    return { data: null, err };
  }
};
/**
 * Returns a string with visual indication of depth in tree.
 * @param {number} depth
 * @returns {string}
 */
const getDepth = (depth: number): string => {
  let sdepth = ''
  if (col > 0) {
    col = 0
    sdepth += '\n'
  }
  for (let x = 0; x < depth - 1; x++) {
    sdepth += ' | '
  }
  sdepth += ' |- '
  return sdepth
}

const createDefaultEmail = () => {
  return {
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
}

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
  let email: Email = createDefaultEmail();  
  try {
    email.from = sender;
    email.to = {
      name: recipients,
      email: msg.displayTo,
    };
    email.text = msg.body;
    email.html = msg.bodyHTML;
    email.subject = msg.subject;
  } catch (err) {
    console.error(err)
  }

  // walk list of attachments and save to fs
  const attachments: Attachment[] = [];
  for (let i = 0; i < msg.numberOfAttachments; i++) {
    const attachment: PSTAttachment = msg.getAttachment(i)
    // Log.debug1(JSON.stringify(activity, null, 2));
    if (attachment.filename) {
      try {
        const attachmentStream = attachment.fileInputStream
        if (attachmentStream) {
          const bufferSize = 8176
          const buffer = Buffer.alloc(bufferSize)
          let bytesRead = 0;
          let mergedBuffer = Buffer.alloc(0);
          do {
            bytesRead = attachmentStream.read(buffer);
            mergedBuffer = Buffer.concat([mergedBuffer, buffer]);
          } while (bytesRead == bufferSize)

          attachments.push({
            name: attachment.filename,
            data: mergedBuffer,
            contentType: attachment.mimeTag
          })
        }
      } catch (err) {
        console.error(err)
      }
    }
  }

  const deliveryDate = msg.messageDeliveryTime || new Date();
  const year = deliveryDate.getFullYear();
  const month = `${deliveryDate.getMonth() + 1}`.padStart(2, "0");
  const day = `${deliveryDate.getDate()}`.padStart(2, "0");
  const uid = msg.descriptorNodeId;
  email.attachments = attachments;
  
  const eml = await buildEml(email);
  if(env.ENV === "dev") {
    const path = `${year}-${month}-${day}-${uid}.eml`;
    fs.writeFileSync(`/Users/tranvinhphuc/scripts/streamPSTFile/output/${path}`, eml);
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
}

/**
 * Get the sender and display.
 * @param {PSTMessage} email
 * @returns {string}
 */
const getSender = (email: PSTMessage): string => {
  let sender = email.senderName
  if (sender !== email.senderEmailAddress) {
    sender += ' (' + email.senderEmailAddress + ')'
  }
  if (verbose && displaySender && email.messageClass === 'IPM.Note') {
    console.log(getDepth(depth) + ' sender: ' + sender)
  }
  return sender
}

/**
 * Get the recipients and display.
 * @param {PSTMessage} email
 * @returns {string}
 */
const getRecipients = (email: PSTMessage): string => {
  // could walk recipients table, but be fast and cheap
  return email.displayTo
}

/**
 * Print a dot representing a message.
 */
const printDot = (): void => {
  process.stdout.write('.')
  if (col++ > 100) {
    console.log('')
    col = 0
  }
}

/**
 * Walk the folder tree recursively and process emails.
 * @param {PSTFolder} folder
 */
const processFolder = async (folder: PSTFolder): Promise<void> => {
  depth++

  // the root folder doesn't have a display name
  if (depth > 0) {
    console.log(getDepth(depth) + folder.displayName)
  }

  // go through the folders...
  if (folder.hasSubfolders) {
    const childFolders: PSTFolder[] = folder.getSubFolders()
    for (const childFolder of childFolders) {
      processFolder(childFolder)
    }
  }

  // and now the emails for this folder
  if (folder.contentCount > 0) {
    depth++
    let email: PSTMessage = folder.getNextChild()
    while (email != null) {
      attempt++;
      if (verbose) {
        console.log(
          getDepth(depth) +
          'Email: ' +
          email.descriptorNodeId +
          ' - ' +
          email.subject
        )
      } else {
        printDot()
      }

      // sender
      const sender = getSender(email)

      // recipients
      const recipients = getRecipients(email)

      // display body?
      if (verbose && displayBody) {
        console.log(highlight('email.body', ANSI_YELLOW), email.body)
        console.log(highlight('email.bodyRTF', ANSI_YELLOW), email.bodyRTF)
        console.log(highlight('email.bodyHTML', ANSI_YELLOW), email.bodyHTML)
      }

      // save content to fs?
      if (saveToFS) {
        // create date string in format YYYY-MM-DD
        let d = email.clientSubmitTime
        if (!d && email.creationTime) {
          d = email.creationTime
        }

        await doSaveToFS(email, sender, recipients);
        if(attempt % (INSERT_SIZE + 1) === INSERT_SIZE) {
          await sleep(SLEEP);
        }
      }
      email = folder.getNextChild()
    }
    depth--
  }
  depth--
}
// load file into memory buffer, then open as PSTFile
const pstFile = new PSTFile(pstFolder)
console.log(pstFile.getMessageStore().displayName)
processFolder(pstFile.getRootFolder()).catch(console.error);