import { request } from ".";

export type Email = {
  object: {
    key: string,
    type: string,
    bucket: string
  }
};

export type Emails_bool_exp = {
  company_id?: {
    _eq: string
  };
  subject?: {
    _in: string[]
  };
  from?: {
    _eq: string
  }
  object?: {
    _cast?: {
      String: {
        _like: string
      }
    }
  }
  _or?: Emails_bool_exp[]
};

export const getEmails = async (condition: Emails_bool_exp): Promise<Email[]> => {
  const gql = /* GraphQL */ `
    query getEmails($condition: emails_bool_exp) {
      emails(where: $condition) {
        object
      }
    }
  `;

  const {
    data: { emails },
  } = await request<{ data: { emails: Email[] } }, { condition: Emails_bool_exp }>(gql, {
    condition,
  });

  return emails;
};