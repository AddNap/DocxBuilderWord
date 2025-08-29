/* global Office, Word */
function insertTileAtSelection(){
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText("{{ FIELD_NAME }}", Word.InsertLocation.replace);
    await context.sync();
  });
}
if (Office.actions){ Office.actions.associate("insertTileAtSelection", insertTileAtSelection); }
