import 'bootstrap/dist/css/bootstrap.css';
import React from 'react';
import ReactDOM from 'react-dom';
import { BrowserRouter } from 'react-router-dom';
import App from './App';
import registerServiceWorker from './registerServiceWorker';
import { initializeIcons } from '@fluentui/react';
import { initializeFileTypeIcons } from '@fluentui/react-file-type-icons';
import { PublicClientApplication  } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { fetchMsalConfig } from "./auth/authConfig";
import * as msTeams from '@microsoft/teams-js';


const baseUrl = document.getElementsByTagName('base')[0].getAttribute('href');
const rootElement = document.getElementById('root');

msTeams.initialize();
initializeIcons("https://res.cdn.office.net/files/fabric-cdn-prod_20241209.001/assets/icons/");
initializeFileTypeIcons("https://res.cdn.office.net/files/fabric-cdn-prod_20241209.001/assets/item-types/");


fetchMsalConfig()
    .then((msalConfig) => {
        const msalInstance = new PublicClientApplication(msalConfig);

        ReactDOM.render(
            <MsalProvider instance={msalInstance}>
                <BrowserRouter basename={baseUrl}>
                    <App />
                </BrowserRouter>
            </MsalProvider>,
            rootElement);
        registerServiceWorker();
    });

