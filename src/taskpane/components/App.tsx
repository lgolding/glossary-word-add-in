import React, { FC, useEffect, useState } from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import Progress from "./Progress";
import GlossaryService from "../services/GlossaryService";

/* global Word, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App: FC<AppProps> = ({ title, isOfficeInitialized }) => {
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

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your add-in to see app body."
      />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo={require("./../../../assets/logo-filled.png")} title={title} message="Welcome" />
      <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={click}>
        Create Glossary Table
      </DefaultButton>
    </div>
  );
};

export default App;
