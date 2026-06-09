import { version as packageVersion } from "../package.json";
import { version as manifestVersion } from "../public/manifest.json";
import { describe, expect, it } from "vitest";

describe("versionsMatch", () => {
  it("manifest.json and package.json should have same version", () => {
    expect(packageVersion).toBe(manifestVersion);
  });
});
