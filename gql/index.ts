import { env } from "process";
import axios from "axios";

const serviceToken: {
  jwt: string | null;
  payload: Record<string, never> | null;
} = {
  jwt: null,
  payload: null,
};

const jwtDecode = (jwt: string) => JSON.parse(Buffer.from(jwt.split(".")[1], "base64").toString());

const isExpiredServiceToken = () => {
  if (!serviceToken.jwt || !serviceToken.payload) {
    return true;
  }
  const bufferTime = 60000; // 1 minute
  return (Date.now() + bufferTime) / 1000 > serviceToken.payload.exp;
};

export const requestServiceToken = async ({ force = false } = {}) => {
  if (force || isExpiredServiceToken()) {
    try {
      const { data } = await axios.post(env.ID_URL, {
        service_id: env.SERVICE_ID,
        secret: env.SERVICE_SECRET,
      });
      serviceToken.jwt = data.token;
      serviceToken.payload = jwtDecode(serviceToken.jwt ?? "");
    } catch (error) {
      console.error(error);
      throw error;
    }
  }
  return serviceToken.jwt;
};

export const request = async <T, K = undefined>(q: string, variables?: K): Promise<T> => {
  if (!q) {
    throw new Error("The graphql query is null");
  }

  const token = await requestServiceToken();
  const { data } = await axios.post(
    `${env.HASURA_ENDPOINT}/v1/graphql`,
    {
      query: q,
      variables,
    },
    {
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${token}`,
      },
    }
  );

  if (data.errors) {
    const message = JSON.stringify(data.errors);
    if (message.includes("invalid-jwt")) {
      console.info(message);
      await requestServiceToken({ force: true });
      return request(q, variables);
    }

    throw new Error(message);
  }
  return data;
};

export type GraphQLResponse<T> = {
  data: T;
  errors: unknown;
};
