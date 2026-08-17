function selectionWouldDeleteWorkbook(state, key) {
  if (key !== "Backspace" && key !== "Delete") return false;
  const { selection, doc } = state;
  if (!selection.empty) {
    let containsWorkbook = false;
    doc.nodesBetween(selection.from, selection.to, (node) => {
      if (node.type.name === "pageWorkbook") containsWorkbook = true;
      return !containsWorkbook;
    });
    return containsWorkbook;
  }

  const $from = selection.$from;
  if ($from.depth !== 1) return false;
  if (key === "Backspace" && $from.parentOffset === 0) {
    return doc.resolve($from.before(1)).nodeBefore?.type?.name === "pageWorkbook";
  }
  if (key === "Delete" && $from.parentOffset === $from.parent.content.size) {
    return doc.resolve($from.after(1)).nodeAfter?.type?.name === "pageWorkbook";
  }
  return false;
}

export { selectionWouldDeleteWorkbook };
