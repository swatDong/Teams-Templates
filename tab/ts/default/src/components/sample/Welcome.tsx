import React from "react";
import { Button, Image } from "@fluentui/react-northstar";
import { Progress } from "./Progress";
import "./Welcome.css";
import { EditCode } from "./EditCode";
import { AzureFunctions } from "./AzureFunctions";
import { Graph } from "./Graph";
import { CurrentUser } from "./CurrentUser";
import { useTeamsFx } from "./lib/useTeamsFx";
import {
  TeamsUserCredential, UserInfo,
} from "teamsdev-client"
import { useData } from "./lib/useData";

export function Welcome(props: { showFunction?: boolean; environment?: string; }) {
  const { showFunction, environment } = {
    showFunction: true,
    environment: "local", // "local" | "azure" | "published"
    ...props,
  };
  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment",
      published: "Teams tenant",
    }[environment] || "local environment";
  const selectedProgressIndex =
    {
      local: 0,
      azure: 1,
      published: 2,
    }[environment] || 0;
  const { isInTeams } = useTeamsFx();
  const credential = new TeamsUserCredential();
  const userProfile = useData(async () => isInTeams ? await credential.getUserInfo() : undefined).data;
  const userName = userProfile ? userProfile.displayName : "";
  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <Image src="thumbsup.png" />
        <h1 className="center">
          Congratulations{userName ? ", " + userName : ""}!
        </h1>
        <p className="center">
          Your app is running in your {friendlyEnvironmentName}
        </p>
        <Progress selectedIndex={selectedProgressIndex}>
          <div>Run in the local environment</div>
          <div>Deploy to the Cloud</div>
          <div>Publish to Teams</div>
        </Progress>
        <div className="sections">
          <EditCode showFunction={showFunction} />
          {isInTeams && <CurrentUser userName={userName} />}
          <Graph />
          {showFunction && <AzureFunctions docsUrl={"https://aka.ms/teamsfx-azure-functions"} />}
        </div>
      </div>
    </div>
  );
}
