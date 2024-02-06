import { env } from "process";
import moment from "moment-timezone";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { uploadContent } from "./s3";
import { getEmails } from "../gql/email";

moment.tz.setDefault("Asia/Ho_Chi_Minh");

export const sleep = (time: number) =>
  new Promise((res) => {
    console.info(`wait for ${time} seconds`);
    setTimeout(() => {
      console.info(`complete wait for ${time} seconds`);
      res(time);
    }, time);
  });

export const getEmlHandler = () => {
  const WORK_DIR =
    env.STAGE === "prod"
      ? path.join(env.COMPANY_ID, env.COMPANY_EMAIL)
      : os.tmpdir();

  const handleEmlDev = async (content, filePath) => {
    const localPath = path.join(WORK_DIR, filePath);
    await fs.ensureDir(path.dirname(localPath));
    await fs.writeFile(localPath, content);
    return localPath;
  };

  const handleEmlProd = async (content, filePath) => {
    const remotePath = path.join(WORK_DIR, filePath);
    await uploadContent({
      remotePath,
      content,
      metaData: {
        group_id: env.GROUP_ID,
        company_id: env.COMPANY_ID,
      },
    });
    return remotePath;
  };

  return env.STAGE === "prod"
    ? { handle: handleEmlProd }
    : { handle: handleEmlDev };
};

export const logger = async (message: string, dest: string) => {
  const now = moment().format("YYYY-MM-DD HH:mm:ss");
  try {
    await fs.appendFile(dest, `${now} - ${message}\n`);
  } catch (err) {
    console.error(`Error writing to log file: ${err}`);
  }
};

export const getEmailsBySubjectsAndSender = async(sender: string, subjects: string[]) => {
  const emails = await getEmails({
    company_id: {
      _eq: env.COMPANY_ID
    },
    subject: {
      _in: subjects
    },
    from: {
      _eq: sender
    }
  });

  return emails;
}