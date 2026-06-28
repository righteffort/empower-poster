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

export interface TokenUpdate {
  type: "TOKEN_UPDATE";
  csrf: string;
}

export interface VisibilityUpdate {
  type: "VISIBILITY_UPDATE";
  show: boolean;
}
