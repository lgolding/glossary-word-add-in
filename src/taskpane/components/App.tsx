import React, { FC, useEffect, useState } from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import GlossaryService from "../GlossaryService";

/* global Word, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App: FC<AppProps> = ({ title, isOfficeInitialized }) => {
  var [listItems, setListItems] = useState<HeroListItem[]>([]);

  useEffect(() => {
    setListItems([
      {
        icon: "Ribbon",
        primaryText: "Achieve more with Office integration",
      },
      {
        icon: "Unlock",
        primaryText: "Unlock features and functionality",
      },
      {
        icon: "Design",
        primaryText: "Create and visualize like a pro",
      },
    ]);
  }, []);

  const click = async () => {
    return Word.run(async (context) => {
      // Insert a Glossary at the end of the document.
      const glossaryService = new GlossaryService(context);
      glossaryService.ensureGlossaryTable();

      await context.sync();
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
      <HeroList message="Discover what Office Add-ins can do for you today!" items={listItems}>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={click}>
          Create Glossary Table
        </DefaultButton>
      </HeroList>
    </div>
  );
};

export default App;
