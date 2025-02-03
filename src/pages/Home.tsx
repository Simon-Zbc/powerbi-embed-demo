import React, { useState } from 'react';
import { PowerBIEmbed } from 'powerbi-client-react';
import 'powerbi-report-authoring';
import { models, Report, Page } from 'powerbi-client';
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
                    "column": "Country",
                    "table": "financials",
                    "schema": "http://powerbi.com/product/schema#column"
                }
            },
            {
                "role": "Y",
                "dataField": {
                    "aggregationFunction": "CountNonNull",
                    "column": "Country",
                    "table": "financials",
                    "schema": "http://powerbi.com/product/schema#columnAggr"
                }
            }
        ]
    }
]

const Home = () => {
    const [inputValue, setInputValue] = useState<string>(JSON.stringify(defaultJson, null, 2));
    const [report, setReport] = useState<Report>();
    const [page, setPage] = useState<Page>();
    const [isEditMode, setIsEditMode] = useState<boolean>(false);

    const handleCreatePage = async () => {
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
            console.error("No page reference found. Please create or fetch the page first.");
            return;
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

                console.log('layout:', layout);
                const customLayout = {
                    x: layout.x,
                    y: layout.y,
                    width: layout.width,
                    height: layout.height,
                    displayState: { mode: models.VisualContainerDisplayMode.Visible }
                };
                const response = await page.createVisual(visualType, customLayout, false);
                const visual = response.visual;

                dataRoles.forEach((dataRole: any) => {
                    visual.addDataField(dataRole.role, dataRole.dataField);
                });
            };
        } catch (error) {
            console.error("Error adding visual:", error);
        }
    };

    const toggleEditMode = () => {
        setIsEditMode(!isEditMode);
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

    return (
        <div className="Home">
            <textarea
                className="Json-textarea"
                placeholder=""
                value={inputValue}
                onChange={(e) => setInputValue(e.target.value)}
            />
            <div className="Button-row">
                <button onClick={handleCreatePage}>Create Page</button>
                <button onClick={handleAddVisual}>Add Visual</button>
                <button onClick={toggleEditMode}>
                    Switch to {isEditMode ? "View" : "Edit"} Mode
                </button>
                <button onClick={handleSave}>Save</button>
            </div>
            <PowerBIEmbed
                embedConfig={{
                    type: 'report',   // Supported types: report, dashboard, tile, visual and qna
                    id: process.env.REACT_APP_EMBED_CONFIG_REPORT_ID, // Report ID
                    embedUrl: 'https://app.powerbi.com/reportEmbed',
                    accessToken: process.env.REACT_APP_EMBED_CONFIG_ACCESS_TOKEN, // Entra ID access token
                    tokenType: models.TokenType.Aad,
                    permissions: models.Permissions.All, // Allow creating/modifying content
                    viewMode: isEditMode ? models.ViewMode.Edit : models.ViewMode.View, // Toggle between Edit and View mode
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
                        ['pageChanged', (event) => console.log(event)],
                    ])
                }

                cssClassName={"Embed-container"}

                getEmbeddedComponent={(embeddedReport) => {
                    setReport(embeddedReport as Report);
                }}
            />
        </div>
    );
};

export default Home;