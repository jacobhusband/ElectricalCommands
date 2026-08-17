import assert from "node:assert/strict";
import test from "node:test";
import { Schema } from "@tiptap/pm/model";
import { EditorState, TextSelection } from "@tiptap/pm/state";
import { selectionWouldDeleteWorkbook } from "./workbookProtection.js";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: { content: "text*", group: "block" },
    pageWorkbook: { group: "block", atom: true },
    text: { group: "inline" },
  },
});

function stateWithSelection(doc, position) {
  return EditorState.create({
    schema,
    doc,
    selection: TextSelection.create(doc, position),
  });
}

test("Backspace at the start of the line after a workbook is protected", () => {
  const doc = schema.node("doc", null, [
    schema.node("pageWorkbook"),
    schema.node("paragraph"),
  ]);
  const state = stateWithSelection(doc, 2);

  assert.equal(selectionWouldDeleteWorkbook(state, "Backspace"), true);
  assert.equal(selectionWouldDeleteWorkbook(state, "Delete"), false);
});

test("Delete at the end of the line before a workbook is protected", () => {
  const doc = schema.node("doc", null, [
    schema.node("paragraph"),
    schema.node("pageWorkbook"),
  ]);
  const state = stateWithSelection(doc, 1);

  assert.equal(selectionWouldDeleteWorkbook(state, "Delete"), true);
  assert.equal(selectionWouldDeleteWorkbook(state, "Backspace"), false);
});

test("ordinary text deletion remains available", () => {
  const doc = schema.node("doc", null, [
    schema.node("pageWorkbook"),
    schema.node("paragraph", null, schema.text("A")),
  ]);
  const state = stateWithSelection(doc, 3);

  assert.equal(selectionWouldDeleteWorkbook(state, "Backspace"), false);
});
