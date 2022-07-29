export default class GlossaryService {
  async ensureGlossaryTable() {
    return Word.run(async (context) => {
      context.document.body.insertTable(2, 2, Word.InsertLocation.end, [["Term", "Definition"]]);
      await context.sync();
    });
  }
}
