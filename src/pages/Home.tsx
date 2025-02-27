import React, { useState, useEffect } from 'react';
import { PowerBIEmbed } from 'powerbi-client-react';
import 'powerbi-report-authoring';
import { models, Report, Page, service, factories } from 'powerbi-client';
import { Configuration, PublicClientApplication } from '@azure/msal-browser';
import '../assets/styles/Home.css';

const defaultJson = [
    {
        "layout": {
            "x": 20,
            "y": 20,
            "width": 400,
            "height": 300
        },
        "visualType": "pieChart",
        "dataRoles": [
            {
                "role": "Category",
                "dataField": {
                    "column": "Column1",
                    "table": "Table1",
                    "schema": "http://powerbi.com/product/schema#column"
                }
            },
            {
                "role": "Y",
                "dataField": {
                    "aggregationFunction": "CountNonNull",
                    "column": "Column1",
                    "table": "Table1",
                    "schema": "http://powerbi.com/product/schema#columnAggr"
                }
            }
        ]
    }
]

const msalConfig: Configuration = {
    auth: {
        clientId: process.env.REACT_APP_CLIENT_ID!,
        authority: process.env.REACT_APP_AUTHORITY,
        redirectUri: process.env.REACT_APP_REDIRECT_URI,
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false,
    }
};

const msalInstance = new PublicClientApplication(msalConfig);

const Home = () => {
    const [inputValue, setInputValue] = useState<string>(JSON.stringify(defaultJson, null, 2));
    const [report, setReport] = useState<Report>();
    const [page, setPage] = useState<Page>();
    const [reportId, setReportId] = useState<string>('');
    const [saveAsName, setSaveAsName] = useState<string>('');
    const [accessToken, setAccessToken] = useState<string>(process.env.REACT_APP_EMBED_CONFIG_ACCESS_TOKEN || '');

    useEffect(() => {
        const initializeMsal = async () => {
            await msalInstance.initialize();
        };
        initializeMsal();
    }, []);

    const handleLogin = async () => {
        try {
            const response = await msalInstance.acquireTokenPopup({
                scopes: ["https://analysis.windows.net/powerbi/api/.default"]
            });
            setAccessToken(response.accessToken);
        } catch (error) {
            console.error("Error login:", error);
        }
    };

    const handleCreateReport = async () => {
        const embedContainer = document.getElementById('embedContainer');
        if (!embedContainer) {
            console.error("Embed container not found.");
            return;
        }

        const groupId = process.env.REACT_APP_EMBED_CONFIG_GROUP_ID;
        const datasetName = process.env.REACT_APP_EMBED_CONFIG_DATASET_NAME;

        try {
            const datasetsResponse = await fetch(`https://api.powerbi.com/v1.0/myorg/groups/${groupId}/datasets`, {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${accessToken}`
                }
            });

            if (!datasetsResponse.ok) {
                if (datasetsResponse.status === 404) {
                    console.error("PowerBI entity not found. Please check the group ID and dataset name.");
                } else {
                    console.error(`Error fetching datasets: ${datasetsResponse.statusText}`);
                }
                return;
            }

            const datasets = await datasetsResponse.json();

            if (datasets.error) {
                console.error(datasets.error.code, datasets.error.message);
                return;
            }

            const dataset = datasets.value.find((ds: any) => ds.name === datasetName);

            if (!dataset) {
                console.error(`Dataset with name ${datasetName} not found`);
                return;
            }

            const embedConfig = {
                type: 'report',
                accessToken: accessToken,
                embedUrl: 'https://app.powerbi.com/reportEmbed',
                datasetId: dataset.id,
                groupId: groupId,
                tokenType: models.TokenType.Aad,
                permissions: models.Permissions.All,
                viewMode: models.ViewMode.Edit,
                settings: {
                    panes: {
                        filters: {
                            expanded: false,
                            visible: false
                        }
                    },
                    background: models.BackgroundType.Default,
                }
            };

            const powerbiService = new service.Service(
                factories.hpmFactory,
                factories.wpmpFactory,
                factories.routerFactory
            );

            const createdReport = powerbiService.createReport(embedContainer, embedConfig);
            setReport(createdReport as Report);
        } catch (error) {
            console.error("Error creating report:", error);
        }
    };

    const handleSaveAs = async () => {
        if (!report) {
            console.error("Report is not initialized.");
            return;
        }

        const saveAsName = "New Report " + new Date().toISOString();

        try {
            await report.saveAs({ name: saveAsName });
            setSaveAsName(saveAsName);
        } catch (error) {
            console.error("Error saving report:", error);
        }
    };

    const handleGetReport = async () => {
        const groupId = process.env.REACT_APP_EMBED_CONFIG_GROUP_ID;

        try {
            const reportsResponse = await fetch(`https://api.powerbi.com/v1.0/myorg/groups/${groupId}/reports`, {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${accessToken}`
                }
            });

            if (!reportsResponse.ok) {
                console.error(`Error fetching reports: ${reportsResponse.statusText}`);
                return;
            }

            const reports = await reportsResponse.json();
            const foundReport = reports.value.find((r: any) => r.name === saveAsName);

            if (!foundReport) {
                console.error(`Report with name ${saveAsName} not found`);
                return;
            }

            setReportId(foundReport.id);
        } catch (error) {
            console.error("Error getting report:", error);
        }
    };

    const handleAddPage = async () => {
        if (!report) {
            console.error("Report is not initialized.");
            return;
        }
        const newPageName = "New Page " + new Date().toISOString();
        try {
            let newPage = await report.addPage(newPageName);
            setPage(newPage);
        } catch (error) {
            console.error("Error adding page:", error);
        }
    };

    const handleAddVisual = async () => {
        if (!page) {
            if (report) {
                const pages = await report.getPages();
                setPage(pages[0]);
            } else {
                console.error("No page reference found. Please create or fetch the page first.");
                return;
            }
        }

        try {
            const visualsJson = JSON.parse(inputValue);
            for (const visualJson of visualsJson) {
                let layout;
                let visualType;
                let dataRoles;

                if (visualJson.layout) {
                    layout = visualJson.layout;
                }
                if (visualJson.visualType) {
                    visualType = visualJson.visualType;
                }
                if (visualJson.dataRoles) {
                    dataRoles = visualJson.dataRoles;
                }

                const customLayout = {
                    x: layout.x,
                    y: layout.y,
                    width: layout.width,
                    height: layout.height,
                    displayState: { mode: models.VisualContainerDisplayMode.Visible }
                };
                if (page) {
                    const response = await page.createVisual(visualType, customLayout, false);
                    const visual = response.visual;
        
                    dataRoles.forEach((dataRole: any) => {
                        visual.addDataField(dataRole.role, dataRole.dataField);
                    });
                }
            };
        } catch (error) {
            console.error("Error adding visual:", error);
        }
    };

    const handleSave = async () => {
        if (!report) {
            console.error("No report reference found. Please create or fetch the report first.");
            return;
        }
        try {
            await report.save();
        } catch (error) {
            console.error("Error saving report:", error);
        }
    };

    const handleExport = async () => {
        if (!report) {
            console.error("No report reference found. Please create or fetch the report first.");
            return;
        }
        try {
            const groupId = process.env.REACT_APP_EMBED_CONFIG_GROUP_ID;
            const exportResponse = await fetch(`https://api.powerbi.com/v1.0/myorg/groups/${groupId}/reports/${reportId}/Export?downloadType=LiveConnect`, {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${accessToken}`
                }
            });

            if (!exportResponse.ok) {
                console.error(exportResponse);
            }

            const blob = await exportResponse.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'report.pbix';
            a.click();
            window.URL.revokeObjectURL(url);
        } catch (error) {
            console.error("Error exporting report:", error);
        }
    };

    return (
        <div className="Home">
            <textarea
                className="Json-textarea"
                placeholder=""
                value={inputValue}
                onChange={(e) => setInputValue(e.target.value)}
            />
            <div className="Button-row">
                <button onClick={handleLogin}>Login</button>
                <button onClick={handleCreateReport}>Create Report</button>
                <button onClick={handleSaveAs}>Save As</button>
                <button onClick={handleGetReport}>Get Report</button>
                <button onClick={handleAddPage}>Add Page</button>
                <button onClick={handleAddVisual}>Add Visual</button>
                <button onClick={handleSave}>Save</button>
                <button onClick={handleExport}>Export</button>
            </div>
            <div id="embedContainer" className="Embed-container">
                {reportId && (
                    <PowerBIEmbed
                        embedConfig={{
                            type: 'report',   // Supported types: report, dashboard, tile, visual and qna
                            id: reportId, // Report ID
                            embedUrl: 'https://app.powerbi.com/reportEmbed',
                            accessToken: accessToken, // Entra ID access token
                            tokenType: models.TokenType.Aad,
                            permissions: models.Permissions.All, // Allow creating/modifying content
                            viewMode: models.ViewMode.Edit,// Toggle between Edit and View mode
                            settings: {
                                panes: {
                                    filters: {
                                        expanded: false,
                                        visible: false
                                    }
                                },
                                background: models.BackgroundType.Default,
                            }
                        }}

                        eventHandlers={
                            new Map([
                                ['loaded', function () { console.log('Report loaded'); }],
                                ['rendered', function () { console.log('Report rendered'); }],
                                ['error', function (event) { console.log(event?.detail); }],
                                ['visualClicked', () => console.log('visual clicked')],
                                ['pageChanged', (event) => {
                                    setPage(event?.detail.newPage);
                                }],
                            ])
                        }

                        cssClassName={"Embed-container"}

                        getEmbeddedComponent={(embeddedReport) => {
                            setReport(embeddedReport as Report);
                        }}
                    />
                )}
            </div>
        </div>
    );
};

export default Home;