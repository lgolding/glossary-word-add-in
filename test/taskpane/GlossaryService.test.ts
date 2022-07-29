const OfficeAddinMock = require("office-addin-mock");
import GlossaryService from "../../src/taskpane/services/GlossaryService";

// Create the seed mock object.
const mockData = {
  context: {
    document: {
      body: {
        insertTable(
          _rowCount: number,
          _columnCount: number,
          _insertLocation: Word.InsertLocation,
          _values: string[][]
        ) {},
      },
    },
  },
  // Mock the Word.InsertLocation enum.
  InsertLocation: {
    end: "end",
  },
  // Mock the Word.run function.
  run: async function (callback: any) {
    await callback(this.context);
  },
};

// Create the mock object from the seed object.
const wordMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called from the GlossaryService functions.
global.Word = wordMock;

// Implement the tests below this line.

describe("The GlossaryService", () => {
  test("should create a table", () => {
    const glossaryService = new GlossaryService();
    glossaryService.ensureGlossaryTable();
  });
});
