export default class GlossaryService {
  constructor(private context: Word.RequestContext) {}

  async ensureGlossaryTable() {
    const body: Word.Body = this.context.document.body;
    const tables = body.tables;
    tables.load("items");
    await this.context.sync();

    const glossaryTable: Word.Table | undefined = this.findGlossaryTable(tables.items);
    if (glossaryTable === undefined) {
      body.insertTable(2, 2, Word.InsertLocation.end, [["Term", "Definition"]]);
    }
  }

  private findGlossaryTable(tableItems: Word.Table[]): Word.Table | undefined {
    if (tableItems.length === 0) {
      return undefined;
    }

    return tableItems[0];
  }
}
