import { describe, expect, it } from "vitest";
import { classifyMimeType } from "./mime";

describe("classifyMimeType", () => {
  it.each([
    ["application/vnd.google-apps.document", "google-doc"],
    ["application/vnd.google-apps.spreadsheet", "google-sheet"],
    ["application/vnd.google-apps.presentation", "google-slides"],
    ["application/vnd.google-apps.shortcut", "shortcut"],
    ["application/vnd.openxmlformats-officedocument.wordprocessingml.document", "docx"],
    ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx"],
    ["application/vnd.openxmlformats-officedocument.presentationml.presentation", "pptx"],
    ["application/pdf", "pdf"],
    ["image/png", "image"],
    ["image/jpeg", "image"],
  ] as const)("%s ⇒ %s", (mimeType, expected) => {
    expect(classifyMimeType(mimeType)).toBe(expected);
  });

  it.each([
    "application/vnd.google-apps.form",
    "application/vnd.google-apps.drawing",
    "application/vnd.google-apps.site",
    "video/mp4",
    "application/zip",
    "text/plain",
  ])("%s ⇒ unsupported", (mimeType) => {
    expect(classifyMimeType(mimeType)).toBe("unsupported");
  });
});
