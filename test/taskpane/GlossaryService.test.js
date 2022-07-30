const OfficeAddinMock = require("office-addin-mock");
import GlossaryService from "../../src/taskpane/services/GlossaryService";

// Create the seed mock object.
const mockData = {
  context: {
    document: {
      body: {
        tables: {
          count: 0,
        },
        // Mock the Body.insertTable method.
        insertTable(_rowCount, _columnCount, _insertLocation, _values) {
          this.tables.count = 1;
        },
      },
    },
  },
  InsertLocation: {
    end: "end",
  },
  run: async function (callback) {
    await callback(this.context);
  },
};

// Create the mock object from the seed object.
const wordMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called from the GlossaryService functions.
global.Word = wordMock;

// Implement the tests below this line.

describe("The GlossaryService", () => {
  test("should create a table", async () => {
    await Word.run(async (context) => {
      // Insert a Glossary at the end of the document.
      const glossaryService = new GlossaryService();
      glossaryService.ensureGlossaryTable(context);

      await context.sync();
    });

    expect(mockData.context.document.body.tables.count).toBe(1);
  });
});
