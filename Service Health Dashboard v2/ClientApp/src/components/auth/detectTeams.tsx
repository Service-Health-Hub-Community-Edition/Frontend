import * as microsoftTeams from "@microsoft/teams-js";

export const checkInTeams = (): boolean => {
    // eslint-disable-next-line dot-notation
    const microsoftTeamsLib = microsoftTeams || window["microsoftTeams"];

    if (!microsoftTeamsLib) {
        return false; // the Microsoft Teams library is for some reason not loaded
    }

    if ((window.parent === window.self && (window as any).nativeInterface) ||
        window.name === "embedded-page-container" ||
        window.name === "extension-tab-frame") {
        return true;
    }
    return false;
};

export const getTeamsTheme = (): string | undefined => {
    // eslint-disable-next-line dot-notation
    const microsoftTeamsLib = microsoftTeams || window["microsoftTeams"];

    if (!microsoftTeamsLib) {
        return undefined;
    }

    var theme: string | undefined;
    microsoftTeams.getContext((context: microsoftTeams.Context) => {
        theme = context.theme;
    });

    return theme;
};