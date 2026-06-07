export interface TokenRequest {
  type: string;
}

export interface TokenResponse {
  csrf: string;
}

export interface PostDataRequest {
  type: string;
  data: any; // eslint-disable-line @typescript-eslint/no-explicit-any
}

export interface PostDataResponse {
  ok: boolean;
  message?: string;
}
