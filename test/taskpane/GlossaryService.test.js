const OfficeAddinMock = require("office-addin-mock");
import GlossaryService from "../../src/taskpane/services/GlossaryService";

var numTables = 0;

// Create the seed mock object.
const mockData = {
  context: {
    document: {
      body: {
        tables: {
          items: [],
        },
        // Mock the Body.insertTable method.
        insertTable(_rowCount, _columnCount, _insertLocation, _values) {
          ++numTables;
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
  test("should create a table!", async () => {
    numTables = 0;

    await Word.run(async (context) => {
      // Insert a Glossary at the end of the document.
      const glossaryService = new GlossaryService(context);
      await glossaryService.ensureGlossaryTable();

      await context.sync();
    });

    expect(numTables).toBe(1);
  });
});
