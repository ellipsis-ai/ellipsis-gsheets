import {google} from 'googleapis';

export interface EllipsisObjectWithEnvVars {
  env: {
    GOOGLE_SERVICE_ACCOUNT_EMAIL: string | undefined,
    GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY: string | undefined
  }
}

export function Client(
  ellipsis: EllipsisObjectWithEnvVars,
  overrideServiceAccountEmail?: string | null,
  overridePrivateKey?: string | null
) {
  const serviceAccountEmail = overrideServiceAccountEmail || ellipsis.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const privateKey = overridePrivateKey || ellipsis.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY;
  if (!serviceAccountEmail || !privateKey) {
    throw new Error("Must provide a service account email address and private key, or set them on the ellipsis env");
  }
  const scope = 'https://www.googleapis.com/auth/drive';
  return new google.auth.JWT({
    email: serviceAccountEmail,
    key: privateKey,
    scopes: [scope],
    subject: serviceAccountEmail
  });
};

export default Client;
