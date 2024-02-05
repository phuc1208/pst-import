import { env } from "process";
import { S3 } from "@aws-sdk/client-s3";
import type { ResponseMetadata } from "@aws-sdk/types";

type ApiResponse<T = any, U = Error> = { data: T; err: U };

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
