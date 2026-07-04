import { describe, it, expect } from "vitest";

import { placeholder } from "../src/app.js";

describe("placeholder", () => {
  it("should do nothing", () => {
    expect(placeholder()).toStrictEqual(true);
  });
});
