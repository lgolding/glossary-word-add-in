export default class GlossaryService {
  private body: Word.Body;
  private tables: Word.TableCollection;

  constructor(private context: Word.RequestContext) {
    this.body = context.document.body;
    this.tables = this.body.tables;
  }

  async ensureGlossaryTable(): Promise<void> {
    this.tables.load("items");
    await this.context.sync();

    const glossaryTable: Word.Table | undefined = this.findGlossaryTable();
    if (glossaryTable === undefined) {
      this.insertGlossaryTable();
    }
  }

  private insertGlossaryTable(): void {
    this.body.insertTable(2, 2, Word.InsertLocation.end, [["Term", "Definition"]]);
  }

  private findGlossaryTable(): Word.Table | undefined {
    if (this.tables.items.length === 0) {
      return undefined;
    }

    return this.tables.items[0];
  }
}
