export interface HoldingEntry {
  /** cusip is typically blank for manual holdings, non-blank otherwise. */
  cusip: string;
  userAccountId: number;
  ticker: string;
  price: number;
  quantity: number;
  value: number;
  fundFees?: number;
}

/**
  A single classification.
  class[0]: top-level asset class, e.g. "U.S. bonds"
  class[1]: asset subclass, e.g. "Government". May be empty.
 */
export interface Classification {
  classes: [string, string];
  fraction: number;
}

/**
  A user account.
 */
export interface Account {
  /**
   The account id from Empower; referenced from `HoldingEntry.userAccountId`.
   array.
   */
  id: number;
  name: string;
}

/**
  A set of Classification objects keyed by ticker. The fractions for one
  ticker's Classification array add to 1.
 */
export type Classifications = Record<string, Classification[]>;

export interface PostPayload {
  version: { major: 0; minor: 5 };
  holdings: HoldingEntry[];
  classifications: Classifications;
  accounts: Account[];
}

export interface SuccessResponse {
  success: true;
  message?: string;
}

export interface ErrorResponse {
  success: false;
  error: string;
}

export type PostResponse = SuccessResponse | ErrorResponse;
