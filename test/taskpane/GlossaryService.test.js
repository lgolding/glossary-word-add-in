const OfficeAddinMock = require("office-addin-mock");
import GlossaryService from "../../src/taskpane/services/GlossaryService";

// Create the seed mock object.
const mockData = {
  context: {
    document: {
      body: {
        tables: {
          items: [],
        },
        // Mock the Body.insertTable method.
        insertTable(rowCount, columnCount, insertLocation, values) {
          this.tables.items.push({
            rowCount,
            columnCount,
            insertLocation,
            values,
          });
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
  test("should create glossary table if it does not already exist", async () => {
    await Word.run(async (context) => {
      const body = context.document.body;
      const tables = body.tables;
      tables.load("items");
      await context.sync();

      // Ensure that there are no tables at all.
      tables.items.length = 0;

      // Insert a Glossary at the end of the document.
      const glossaryService = new GlossaryService(context);
      await glossaryService.ensureGlossaryTable();

      await context.sync();
    });

    expect(wordMock.context.document.body.tables.items.length).toBe(1);
  });

  test("should not create glossary table if it already exists", async () => {
    await Word.run(async (context) => {
      const body = context.document.body;
      const tables = body.tables;
      tables.load("items");
      await context.sync();

      // Ensure that there is already a table.
      tables.items.length = 0;
      body.insertTable(2, 2, "End", [[]]);

      // Insert a Glossary at the end of the document.
      const glossaryService = new GlossaryService(context);
      await glossaryService.ensureGlossaryTable();

      await context.sync();
    });

    expect(wordMock.context.document.body.tables.items.length).toBe(1);
  });
});
