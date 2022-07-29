export default class GlossaryService {
  constructor(private context: Word.RequestContext) {}

  ensureGlossaryTable() {
    this.context.document.body.insertTable(2, 2, Word.InsertLocation.end, [["Term", "Definition"]]);
  }
}
