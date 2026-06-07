import { type FileHandle, open, readFile } from "fs/promises";
import {
  getClassifications,
  getHoldings,
} from "../src/entrypoints/content/processing";

async function readJsonFile<T>(infile: string | FileHandle): Promise<T> {
  try {
    const fileContent = await readFile(infile, "utf-8");
    return JSON.parse(fileContent);
  } catch (error) {
    if (error instanceof Error) {
      throw new Error(`Failed to read JSON file: ${error.message}`);
    }
    throw error;
  }
}

async function main() {
  const spData = (await readJsonFile(await open("/dev/stdin", "r"))) as any;
  const classificationsIn = spData.classifications[0].classifications;
  const holdingsIn = getHoldings(spData.holdings);
  const holdings = getHoldings(holdingsIn);
  const classifications = getClassifications(classificationsIn);
  console.log(JSON.stringify({ holdings, classifications }));
}
await main();
