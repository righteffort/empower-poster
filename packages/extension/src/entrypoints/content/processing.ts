import type {
  HoldingEntry,
  Classifications,
} from "@righteffort/empower-poster-types";

export type { Classifications };

export interface ClassificationError {
  ticker: string;
  classes: [string, string];
  fraction: number;
}

export interface ClassificationsResult {
  classifications: Classifications;
  errors: ClassificationError[];
}

interface ClassificationErrorInternal {
  ticker: string;
  classes: [string, string];
  value: number;
  userAccountId: number;
}

// Holding data returned by Empower
interface HoldingEntryIn {
  cusip?: string;
  userAccountId: number;
  description?: string;
  ticker?: string;
  price: number;
  quantity: number;
  value: number;
}

// Classification data returned by Empower. Exported for tests.
export interface ClassificationIn {
  classificationTypeName: string;
  assets?: Asset[];
  classifications?: ClassificationIn[];
}

interface Asset {
  ticker?: string;
  description?: string;
  value: number;
  userAccountId: number;
}

// Internal representation of one classification row by value
interface ClassificationValue {
  ticker: string;
  classes: [string, string];
  value: number;
  userAccountId: number;
}

function cleanTicker(ticker: string): string {
  // For whatever reason FSPSX can show up at Empower as FSIVX.
  if (ticker === "FSIVX") {
    return "FSPSX";
  }
  return ticker;
}

export function getHoldings(holdingsIn: HoldingEntryIn[]): HoldingEntry[] {
  return holdingsIn
    .map((h) => ({
      cusip: h.cusip || "",
      userAccountId: h.userAccountId,
      price: h.price,
      quantity: h.quantity,
      value: h.value,
      ticker: cleanTicker(h.ticker || "") || h.description || "",
    }))
    .filter((h) => h.quantity !== 0);
}

interface ProcessClassificationsResult {
  values: ClassificationValue[];
  errors: ClassificationErrorInternal[];
}

// Flatten Empower classification data. Also return details for any negative allocations.
function processClassifications(
  input: ClassificationIn[],
): ProcessClassificationsResult {
  const valueMap: ClassificationValue[] = [];
  const errors: ClassificationErrorInternal[] = [];

  function processAssets(
    assets: Asset[],
    className: string,
    subClassName: string,
  ) {
    for (const asset of assets) {
      const ticker = cleanTicker(asset.ticker || "") || asset.description || "";
      if (asset.value < 0) {
        errors.push({
          ticker,
          classes: [className, subClassName],
          value: asset.value,
          userAccountId: asset.userAccountId,
        });
      }
      const adjustedValue = Math.max(0, asset.value);
      valueMap.push({
        ticker,
        classes: [className, subClassName],
        value: adjustedValue,
        userAccountId: asset.userAccountId,
      });
    }
  }

  for (const item of input) {
    const className = item.classificationTypeName || "";
    // Process top-level assets
    if (item.assets) {
      processAssets(item.assets, className, "");
    }
    // Process nested classifications
    if (item.classifications) {
      for (const subClassification of item.classifications) {
        const subClassName = subClassification.classificationTypeName || "";
        if (subClassification.assets) {
          processAssets(subClassification.assets, className, subClassName);
        }
      }
    }
  }
  return { values: valueMap, errors };
}

const PRECISION = 100000000000;

function groupByTickerWithFractions(
  entries: ClassificationValue[],
  errors: ClassificationErrorInternal[],
): ClassificationsResult {
  const tickerAccountTotals = new Map<string, number>();

  // Group and aggregate by ticker and classification
  const tickerMap = new Map<string, Map<string, number>>();
  for (const entry of entries) {
    const classKey = entry.classes.join("\0");
    const tickerClassMap =
      tickerMap.get(entry.ticker) ?? new Map<string, number>();
    if (!tickerMap.has(entry.ticker))
      tickerMap.set(entry.ticker, tickerClassMap);

    tickerClassMap.set(
      classKey,
      (tickerClassMap.get(classKey) ?? 0) + entry.value,
    );
  }

  // Calculate fractions and transform
  const classifications = Object.fromEntries(
    Array.from(tickerMap, ([ticker, classMap]) => {
      const total = Array.from(classMap.values()).reduce(
        (sum, value) => sum + value,
        0,
      );
      return [
        ticker,
        Array.from(classMap, ([classKey, value]) => ({
          classes: classKey.split("\0") as [string, string],
          // Treat 0/0 as 1
          fraction:
            value == total
              ? 1
              : Math.round((value / total) * PRECISION) / PRECISION,
        })).filter((e) => e.fraction !== 0),
      ];
    }),
  );

  // Calculate error percentages
  let processedErrors: ClassificationError[] = [];
  if (errors.length > 0) {
    // Build ticker account totals in one pass
    for (const entry of [...entries, ...errors]) {
      const tickerAccountKey = `${entry.ticker}\0${entry.userAccountId}`;
      tickerAccountTotals.set(
        tickerAccountKey,
        (tickerAccountTotals.get(tickerAccountKey) ?? 0) + entry.value,
      );
    }

    processedErrors = errors.map((error) => {
      const tickerAccountKey = `${error.ticker}\0${error.userAccountId}`;
      const tickerAccountTotal = tickerAccountTotals.get(tickerAccountKey) || 0;
      return {
        ticker: error.ticker,
        classes: error.classes,
        fraction:
          tickerAccountTotal !== 0 ? error.value / tickerAccountTotal : -1,
      };
    });
  }

  return { classifications, errors: processedErrors };
}

export function getClassifications(
  input: ClassificationIn[],
): ClassificationsResult {
  const { values: adjustedEntries, errors } = processClassifications(input);
  return groupByTickerWithFractions(adjustedEntries, errors);
}
