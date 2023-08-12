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

        // ---- ADDED: set slicer value -----
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
                  //this need to be stored and retrieved from local storage
                  value: "2023-07-04T00:00:00.000"
                },
                {
                  operator: "LessThan",
                  //this need to be stored and retrieved from local storage
                  //the date is always 1 day after the selected date
                  value: "2023-07-07T00:00:00.000"
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

            // ---- ADDED pass slicer
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
            
                // Retrieve all visuals with the type "slicer".
                let slicers = visuals.filter(function (visual) {
                    return visual.type === "slicer";
                });
            
                slicers.forEach(async (slicer) => {
                    // Get the slicer state.
                    const state = await slicer.getSlicerState();
                    // console.log(state['targets'][0]['table'])
                    if (state['targets'][0]['table'] === 'Serve Date' && state['targets'][0]['column'] === 'Date Code') {
                        slicer.on('selectionChange', function(e) {
                            localStorage.setItem('fromDate', state['filters'][0]['conditions'][0]['value'])
                            localStorage.setItem('toDate', state['filters'][0]['conditions'][1]['value'])
                        })
                    }
                });
                return slicerValues
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