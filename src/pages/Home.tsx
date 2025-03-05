import React, { useState, useEffect } from 'react';
import { PowerBIEmbed } from 'powerbi-client-react';
import 'powerbi-report-authoring';
import { models, Report, Page, service, factories } from 'powerbi-client';
import { Configuration, PublicClientApplication } from '@azure/msal-browser';
import '../assets/styles/Home.css';

const defaultJson = {
    "report": {
        "title": "Sample Report",
        "pages": [
            {
                "title": "Sample Page 1",
                "visuals": [
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
                    },
                    {
                        "layout": {
                            "x": 500,
                            "y": 20,
                            "width": 400,
                            "height": 300
                        },
                        "visualType": "columnChart",
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
                                "role": "Series",
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
            },
            {
                "title": "Sample Page 2",
                "visuals": [
                    {
                        "layout": {
                            "x": 20,
                            "y": 20,
                            "width": 400,
                            "height": 300
                        },
                        "visualType": "columnChart",
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
                                "role": "Series",
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
                    },
                    {
                        "layout": {
                            "x": 500,
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
            }
        ]
    }
}

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
            try {
                await msalInstance.initialize();
            } catch (error) {
                alert("Error initializing MSAL: " + (error as any).detailedMessage);
            }
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
            alert("Error login: " + (error as any).detailedMessage);
        }
    };

    const handleCreateReport = async () => {
        const embedContainer = document.getElementById('embedContainer');
        if (!embedContainer) {
            alert("Embed container not found.");
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
                alert(`Error fetching datasets: ${datasetsResponse.status}. Please check the group ID and dataset name.`);
                return;
            }

            const datasets = await datasetsResponse.json();

            if (datasets.error) {
                alert(datasets.error.code + ": " + datasets.error.message);
                return;
            }

            const dataset = datasets.value.find((ds: any) => ds.name === datasetName);

            if (!dataset) {
                alert(`Dataset with name ${datasetName} not found`);
                return;
            }

            const reportJson = JSON.parse(inputValue);
            const embedConfig = {
                type: 'report',
                accessToken: accessToken,
                embedUrl: 'https://app.powerbi.com/reportEmbed',
                datasetId: dataset.id,
                groupId: groupId,
                tokenType: models.TokenType.Aad,
                pageName: reportJson.report.pages[0].title,
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
            alert("Error creating report: " + (error as any).detailedMessage);
        }
    };

    const handleSaveAs = async () => {
        if (!report) {
            alert("Report is not initialized.");
            return;
        }

        try {
            const reportJson = JSON.parse(inputValue);
            await report.saveAs({ name: reportJson.report.title });
            setSaveAsName(reportJson.report.title);
        } catch (error) {
            alert("Error saving report: " + (error as any).detailedMessage);
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
                alert(`Error fetching reports: ${reportsResponse.status}`);
                return;
            }

            const reports = await reportsResponse.json();
            const foundReport = reports.value.find((r: any) => r.name === saveAsName);

            if (!foundReport) {
                alert(`Report with name ${saveAsName} not found`);
                return;
            }

            setReportId(foundReport.id);
        } catch (error) {
            alert("Error getting report: " + (error as any).detailedMessage);
        }
    };

    const handleCreateVisual = async () => {
        let page: Page | undefined;

        const reportJson = JSON.parse(inputValue);
        for (let index = 0; index < reportJson.report.pages.length; index++) {
            if (report) {
                try {
                    const pages = await report.getPages();
                    if (index === 0) {
                        page = pages[0];
                        await report.renamePage(pages[0].name, reportJson.report.pages[index].title);
                    } else {
                        page = await report.addPage(reportJson.report.pages[index].title);
                    }
                    await report.setPage(page.name);
                } catch (error) {
                    alert("Error creating page: " + (error as any).detailedMessage);
                    return;
                }
            }

            for (const visualJson of reportJson.report.pages[index].visuals) {
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

                try {
                    if (page) {
                        const response = await page.createVisual(visualType, customLayout, false);
                        const visual = response.visual;
                        // const capabilities = await visual.getCapabilities();
                        // console.log(visualType, capabilities);
                        for (const dataRole of dataRoles) {
                            await visual.addDataField(dataRole.role, dataRole.dataField);
                        }
                    }
                } catch (error) {
                    alert("Error creating visual: " + (error as any).detailedMessage);
                }

            };
        }
    };

    const handleSave = async () => {
        if (!report) {
            alert("No report reference found. Please create or fetch the report first.");
            return;
        }
        try {
            await report.save();
        } catch (error) {
            alert("Error saving report: " + (error as any).detailedMessage);
        }
    };

    const handleExport = async () => {
        if (!report) {
            alert("No report reference found. Please create or fetch the report first.");
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
                alert(`Error exporting report: ${exportResponse.status}`);
                return;
            }

            const blob = await exportResponse.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'report.pbix';
            a.click();
            window.URL.revokeObjectURL(url);
        } catch (error) {
            alert("Error exporting report: " + (error as any).detailedMessage);
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
                <button onClick={handleCreateVisual}>Create Visual</button>
                <button onClick={handleSave}>Save</button>
                <button onClick={handleExport}>Export</button>
                {/* <button onClick={handleTest}>Test</button> */}
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