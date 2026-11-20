import { describe, expect, it } from "vitest";
import { SOURCE_TYPE_BY_EXTENSION } from "./mime-types";

// The server reads the type from this table when a browser declares none; the upload screen's admission list is held equal to it by `core/upload-file-types.test.ts`.
describe("SOURCE_TYPE_BY_EXTENSION", () => {
  it("🔴 jpeg is a JPEG, like jpg (a photo.jpeg with no declared type was refused as UNSUPPORTED_TYPE)", () => {
    expect(SOURCE_TYPE_BY_EXTENSION["jpeg"]).toBe("image/jpeg");
    expect(SOURCE_TYPE_BY_EXTENSION["jpg"]).toBe("image/jpeg");
  });
});
