export default class GlossaryService {
  ensureGlossaryTable(context: Word.RequestContext) {
    context.document.body.insertTable(2, 2, Word.InsertLocation.end, [["Term", "Definition"]]);
  }
}
