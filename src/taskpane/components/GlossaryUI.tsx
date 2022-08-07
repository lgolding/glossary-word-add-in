import React, { FC } from "react";
import { DefaultButton } from "@fluentui/react";
import GlossaryService from "../services/GlossaryService";

interface GlossaryUIProps {}

const GlossaryUI: FC<GlossaryUIProps> = ({}) => {
  const click = async () => {
    return Word.run(async (context: Word.RequestContext) => {
      // Insert a Glossary at the end of the document.
      const glossaryService = new GlossaryService(context);
      await glossaryService.ensureGlossaryTable();

      await context.sync();
    }).catch(function (error: any) {
      // Catch and log any errors that occur within `Word.run`.
      console.log(`Error: ${error}`);
      if (error instanceof OfficeExtension.Error) {
        console.log(`Debug information: ${JSON.stringify(error.debugInfo)}`);
      }
    });
  };

  return (
    <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={click}>
      Create Glossary Table
    </DefaultButton>
  );
};

export default GlossaryUI;
