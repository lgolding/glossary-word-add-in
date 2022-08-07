export default class GlossaryService {
  constructor(private context: Word.RequestContext) {}

  async ensureGlossaryTable() {
    const body: Word.Body = this.context.document.body;
    const tables = body.tables;
    tables.load("items");
    await this.context.sync();

    if (tables.items.length === 0) {
      body.insertTable(2, 2, Word.InsertLocation.end, [["Term", "Definition"]]);
    }
  }
}
