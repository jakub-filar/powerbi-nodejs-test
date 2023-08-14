// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

let models = window["powerbi-client"].models;
let reportContainer = $("#report-container").get(0);

// Initialize iframe for embedding report
powerbi.bootstrap(reportContainer, { type: "report" });

// AJAX request to get the report details from the API and pass it to the UI
$.ajax({
    type: "GET",
    url: "/getEmbedToken",
    dataType: "json",
    success: function (embedData) {

        // ---- -------
        const weekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0] + "T23:00:00.000Z"
        const yesterday = new Date(Date.now() - 1 * 24 * 60 * 60 * 1000).toISOString().split('T')[0] + "T23:00:00.000Z"
        const storedStartDate = localStorage.getItem('returnsDashboardStoredStartDate')
        const storedEndDate = localStorage.getItem('returnsDashboardStoredEndDate')

        const target = {
            table: "Serve Date",
            column: "Date Code"
        }
        
        const dateFilter = {
            $schema: "http://powerbi.com/product/schema#advanced",
            target: target,
            logicalOperator: "And",
            conditions: [
                {
                  operator: "GreaterThanOrEqual",
                  //date in format YYYY-MM-DDT23:00:00.000Z (eg. 2023-07-03T23:00:00.000Z)
                  value: storedStartDate ? storedStartDate : weekAgo
                },
                {
                  operator: "LessThan",
                  //date in format YYYY-MM-DDT23:00:00.000Z (eg. 2023-07-12T23:00:00.000Z)
                  value: storedEndDate ? storedEndDate : yesterday
                }
              ],
            filterType: models.FilterType.Advanced
        }
        
        const slicers = [
            {
              selector: {
                $schema: "http://powerbi.com/product/schema#slicerTargetSelector",
                target: target 
              },
              state: {
                filters: [dateFilter]
              }
            }
        ];
        // ---- -----

        // Create a config object with type of the object, Embed details and Token Type
        let reportLoadConfig = {
            type: "report",
            tokenType: models.TokenType.Embed,
            accessToken: embedData.accessToken,

            // Use other embed report config based on the requirement. We have used the first one for demo purpose
            embedUrl: embedData.embedUrl[0].embedUrl,

            // ---- pass slicer
            slicers: slicers
            // -------

            // Enable this setting to remove gray shoulders from embedded report
            // settings: {
            //     background: models.BackgroundType.Transparent
            // }
        };

        // Use the token expiry to regenerate Embed token for seamless end user experience
        // Refer https://aka.ms/RefreshEmbedToken
        tokenExpiry = embedData.expiry;

        // Embed Power BI report when Access token and Embed URL are available
        let report = powerbi.embed(reportContainer, reportLoadConfig);

        // Clear any other loaded handler events
        report.off("loaded");

        // Triggers when a report schema is successfully loaded
        report.on("loaded", function () {
            console.log("Report load successful");
        });

        // Clear any other rendered handler events
        report.off("rendered");

        // Triggers when a report is successfully embedded in UI
        report.on("rendered", async function () {
            try {
                const pages = await report.getPages();
                // Retrieve the active page.
                let pageWithSlicer = pages.filter(function (page) {
                    return page.isActive;
                })[0];
            
                const visuals = await pageWithSlicer.getVisuals();
            
                // Retrieve all visuals with the type "slicer"
                let slicers = visuals.filter(function (visual) {
                    return visual.type === "slicer";
                });
            
                slicers.forEach(async (slicer) => {
                    // Get the slicer state.
                    const state = await slicer.getSlicerState();
                    
                    // find slicer that filters table Serve Date and column Date Code
                    if (state['targets'][0]['table'] === 'Serve Date' && state['targets'][0]['column'] === 'Date Code') {
                        const slicerStartDate = state['filters'][0]['conditions'][0]['value']
                        const slicerEndDate = state['filters'][0]['conditions'][1]['value']
                        localStorage.setItem('returnsDashboardStoredStartDate', slicerStartDate)
                        localStorage.setItem('returnsDashboardStoredEndDate', slicerEndDate)
                    }
                });
            }
            catch (e) {}
        });

        // Clear any other error handler events
        report.off("error");

        // Handle embed errors
        report.on("error", function (event) {
            let errorMsg = event.detail;
            console.error(errorMsg);
            return;
        });
    },

    error: function (err) {

        // Show error container
        let errorContainer = $(".error-container");
        $(".embed-container").hide();
        errorContainer.show();

        // Get the error message from err object
        let errMsg = JSON.parse(err.responseText)['error'];

        // Split the message with \r\n delimiter to get the errors from the error message
        let errorLines = errMsg.split("\r\n");

        // Create error header
        let errHeader = document.createElement("p");
        let strong = document.createElement("strong");
        let node = document.createTextNode("Error Details:");

        // Get the error container
        let errContainer = errorContainer.get(0);

        // Add the error header in the container
        strong.appendChild(node);
        errHeader.appendChild(strong);
        errContainer.appendChild(errHeader);

        // Create <p> as per the length of the array and append them to the container
        errorLines.forEach(element => {
            let errorContent = document.createElement("p");
            let node = document.createTextNode(element);
            errorContent.appendChild(node);
            errContainer.appendChild(errorContent);
        });
    }
});