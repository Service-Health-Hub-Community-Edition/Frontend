import * as React from 'react';
import { Text } from '@fluentui/react';
import { Dropdown, IDropdownOption, IDropdownStyles } from '@fluentui/react';
import { PrimaryButton } from '@fluentui/react';
import { GlobalState } from './GlobalState';
import { translate, getUIString } from "../translation/translator";
import { acquireAccessToken } from "../auth/AccessTokenHelper";

const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 250 } };

export interface ISettingsState {
    language: string;
    languageOptions: IDropdownOption[];
    uiStrings: Map<string, string>;
    saveSuccess?: boolean;
    error?: string;
}

export class Settings extends React.Component<{}, ISettingsState> {
    static contextType = GlobalState;

    constructor(props: { }) {
        super(props);

        var uiStrings: Map<string, string> = new Map([
            ["Settings", "Settings"],
            ["Language", "Language"],
            ["Select language", "Select language:"],
            ["Save", "Save"],
            ["User settings successfully saved.", "User settings successfully saved."],
            ["Language settings apply for Company dashboard only","Language settings apply for Company dashboard only"]
        ]);

        this.state = {
            language: "",
            languageOptions: [],
            uiStrings: uiStrings,
            saveSuccess: undefined,
            error: undefined
        };
    }

    public render() {
        const {
            language, languageOptions, error, uiStrings, saveSuccess } = this.state;

        return (          
            <div className="userSettings">
                <Text variant={'mediumPlus'} block>
                    <b>{getUIString(uiStrings, "Settings")}</b>
                </Text>
                <Dropdown
                    label={getUIString(uiStrings, "Language")}
                    selectedKey={language ? language : "en"}
                    onChange={this._onLanguageDropdownChange}
                    placeholder={getUIString(uiStrings, "Select language")}
                    options={languageOptions}
                    styles={dropdownStyles}
                    isDisabled={languageOptions.length == 0}
                />
                <Text variant={'small'} block style={{ color: '#101010' }}>
                    <i>{getUIString(uiStrings, "Language settings apply for Company dashboard only")}</i>
                </Text>
                <br />
                <PrimaryButton onClick={this._setUserSettings} text={getUIString(uiStrings, "Save")} />
                
                <Text variant={'small'} block style={{ color: '#ff0000', display: saveSuccess ? "" : "none" }}>
                    <br />
                    {getUIString(uiStrings, "User settings successfully saved.")}
                </Text>
            </div >
        );
    }

    componentDidMount() {
        let globalState: any = this.context;
        var languageList: IDropdownOption[] = [];
        var translatedUIStrings: Map<string, string>;
        var language: string = globalState.getProfileProperty("language");

        this.setState({
            language: language
        });

        acquireAccessToken()
            .then((response) => {
                translate(language,
                    this.state.uiStrings,
                    response
                ).then((response) => {
                    translatedUIStrings = response;
                    this.setState({
                        uiStrings: translatedUIStrings
                    });
                });
            }).then(() => {
        fetch('https://api.cognitive.microsofttranslator.com/languages?api-version=3.0')
            .then(response => {
                if (response.ok) {
                    return response.json();
                } else {
                    this.setState({
                        error: response.status + " " + response.statusText
                    });
                    throw Error(response.status + " " + response.statusText);
                }
            })
            .then((result) => {
                for (const lang in result.translation) {
                    languageList.push({
                        key: lang,
                        text: result.translation[lang].name
                    });
                }

                languageList = languageList.sort((a, b) => 0 - (a.text > b.text ? -1 : 1));

                const selectedLanguage = language;

                this.setState({
                    language: selectedLanguage,
                    languageOptions: languageList
                });
            });
        });
    }

    private _onLanguageDropdownChange = (event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption): void => {
        this.setState({
            language: item ? item.key.toString() : "en",
            saveSuccess: false
        });
    };

    private _setUserSettings = (): void => {
        let globalState: any = this.context;
        var translatedUIStrings: Map<string, string>;
        globalState.setProfileProperty("language", this.state.language);

        this.setState({
            saveSuccess: true
        });

        acquireAccessToken()
            .then((response) => {
                translate(this.state.language,
                    this.state.uiStrings,
                    response
                ).then((response) => {
                    translatedUIStrings = response;
                    this.setState({
                        uiStrings: translatedUIStrings
                    });
                });
            });
    }
}