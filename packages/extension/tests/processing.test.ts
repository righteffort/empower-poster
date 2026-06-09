import { describe, it, expect } from "vitest";
import { getHoldings, getClassifications } from "../src/processing.js";
import type {
  Classifications,
  ClassificationIn,
  ClassificationError,
} from "../src/processing.js";
import { globSync, readFileSync } from "fs";

describe("getHoldings", () => {
  it("should run getHoldings correctly", () => {
    for (const inputFile of globSync(
      "{,private/}testdata/*getHoldings*-input.json",
    )) {
      const expectedFile = inputFile.replace("-input", "-expected");
      const input = JSON.parse(readFileSync(inputFile, "utf-8"));
      const actual = getHoldings(input);
      // const actualFile = inputFile.replace('-input','-actual');
      // writeFileSync(actualFile, JSON.stringify(actual, null, 2));
      const expected = JSON.parse(readFileSync(expectedFile, "utf-8"));
      expect(new Set(actual)).toEqual(new Set(expected));
    }
  });
});

// TODO: move this stuff to a Smallify class in a different file.
// let tickerCount = 0;
// const tickerMap = new Map<string, {tickerOut: string, descriptionOut: string}>();

// function smallifyAsset(input: Asset): Asset {
//   /* eslint-disable no-prototype-builtins */
//   const {ticker, description, value} = input;
//   const valueOut = Math.sign(value) * (Math.round((Math.random()*9999+ 1))/100);
//   let descriptionOut = description;
//   let tickerOut: string | undefined;
//   if (ticker !== undefined) {
//     let td = tickerMap.get(ticker);
//     if (!td) {
//       tickerCount++;
//       td = {tickerOut: `TICK${tickerCount}`, descriptionOut: `Description ${tickerCount}`}
//       tickerMap.set(ticker, td);
//     }
//     // TODO y not {tickerOut, descriptionOut} = td;
//     tickerOut = td.tickerOut;
//     descriptionOut = td.descriptionOut;
//   }
//   if (ticker === undefined) {
//     throw new Error("tickerOut not set, should be unreachable");
//   }
//   return {
//     ...(input.hasOwnProperty('ticker') && {ticker: tickerOut}),
//     ...(input.hasOwnProperty('description') && {description: descriptionOut}),
//     ...(input.hasOwnProperty('value') && {value: valueOut}),
//   } as Asset;
//   /* eslint-enable no-prototype-builtins */
// }

// function smallifyClassification(input: Classification): Classification {
//   /* eslint-disable no-prototype-builtins */
//   return {
//     ...(input.hasOwnProperty('classificationTypeName') && {classificationTypeName: input.classificationTypeName}),
//     ...(input.hasOwnProperty('assets') && {assets: input.assets.map(a => smallifyAsset(a))}),
//     ...(input.hasOwnProperty('classifications') && {classifications: input.classifications.map(c => smallifyClassification(c))}),
//   } as Classification;
//   /* eslint-enable no-prototype-builtins */
// }

function canonicalizeClassifications(
  classifications: Classifications,
): Classifications {
  return Object.fromEntries(
    Object.entries(classifications).map(([ticker, classificationValues]) => [
      ticker,
      classificationValues.sort((a, b) => {
        const keyA = a.classes.join("\0");
        const keyB = b.classes.join("\0");
        return keyA.localeCompare(keyB);
      }),
    ]),
  );

  // return Object.fromEntries(
  //   Array.from(classifications.entries(), ([ticker, classificationValues]) => [
  //     ticker,
  //     classificationValues.sort((a, b) => {
  //       const keyA = a.classes.join("\0");
  //       const keyB = b.classes.join("\0");
  //       return keyA.localeCompare(keyB);
  //     }),
  //   ]),
  // );
}

function canonicalizeErrors(
  errors: ClassificationError[],
): ClassificationError[] {
  return errors.sort((a, b) => {
    const tickerCompare = a.ticker.localeCompare(b.ticker);
    if (tickerCompare !== 0) return tickerCompare;
    const classKeyA = a.classes.join("\0");
    const classKeyB = b.classes.join("\0");
    return classKeyA.localeCompare(classKeyB);
  });
}

describe("getClassifications", () => {
  it("should run getClassifications correctly", () => {
    for (const inputFile of globSync(
      "{,private/}testdata/*getClassifications*-input.json",
    )) {
      const expectedFile = inputFile.replace("-input", "-expected");
      const input = JSON.parse(
        readFileSync(inputFile, "utf-8"),
      ) as ClassificationIn[];
      // const smallFile = inputFile.replace('-input','-small');
      // writeFileSync(smallFile, JSON.stringify(input.map(c => smallifyClassification(c)), null, 2));
      // console.log(`wrote ${smallFile}`);
      const actual = getClassifications(input);
      // const actualFile = inputFile.replace("-input", "-actual");
      // writeFileSync(actualFile, JSON.stringify(actual, null, 2));
      const expected = JSON.parse(readFileSync(expectedFile, "utf-8"));
      expect(canonicalizeClassifications(actual.classifications)).toEqual(
        canonicalizeClassifications(expected.classifications),
      );
      expect(canonicalizeErrors(actual.errors)).toEqual(
        canonicalizeErrors(expected.errors),
      );
    }
  });
});
