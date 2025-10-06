import makeConstructorArgs from "../utilities/commonUtils";
import { Visual } from "../../PowerBI-SPC/src/visual";
import * as d3 from "../../PowerBI-SPC/src/D3 Plotting Functions/D3 Modules"
const dummyData = [{"categorical":{"categories":[{"source":{"roles":{"key":true},"type":{"temporal":{"underlyingType":519}}},"values":["Mon Jan 01 2024 08:00:00 GMT+0800","Thu Feb 01 2024 08:00:00 GMT+0800","Fri Mar 01 2024 08:00:00 GMT+0800","Mon Apr 01 2024 08:00:00 GMT+0800","Wed May 01 2024 08:00:00 GMT+0800","Sat Jun 01 2024 08:00:00 GMT+0800","Mon Jul 01 2024 08:00:00 GMT+0800","Thu Aug 01 2024 08:00:00 GMT+0800","Sun Sep 01 2024 08:00:00 GMT+0800","Tue Oct 01 2024 08:00:00 GMT+0800","Fri Nov 01 2024 08:00:00 GMT+0800","Sun Dec 01 2024 08:00:00 GMT+0800","Wed Jan 01 2025 08:00:00 GMT+0800","Sat Feb 01 2025 08:00:00 GMT+0800","Sat Mar 01 2025 08:00:00 GMT+0800","Tue Apr 01 2025 08:00:00 GMT+0800","Thu May 01 2025 08:00:00 GMT+0800","Sun Jun 01 2025 08:00:00 GMT+0800","Tue Jul 01 2025 08:00:00 GMT+0800","Fri Aug 01 2025 08:00:00 GMT+0800","Mon Sep 01 2025 08:00:00 GMT+0800","Wed Oct 01 2025 08:00:00 GMT+0800","Sat Nov 01 2025 08:00:00 GMT+0800","Mon Dec 01 2025 08:00:00 GMT+0800"],"objects":[{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}},{"canvas":{"show_errors":true,"lower_padding":60,"upper_padding":10,"left_padding":60,"right_padding":10},"spc":{},"outliers":{},"nhs_icons":{},"scatter":{},"lines":{},"x_axis":{},"y_axis":{},"dates":{},"labels":{}}]}],"values":[{"source":{"roles":{"numerators":true}},"values":[0.0673,0.7948,1.5177,-2.1649,-0.2144,0.1286,0.2253,-2.1378,-2.056,-2.0566,-0.0523,-0.7576,-0.0093,0.7548,0.7776,0.5302,-0.2699,-0.5413,0.6965,0.3852,-0.8818,-1.525,0.5477,2.944]}]}}]

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("create-plot").onclick = () => tryCatch(createPlot);
  }
});

async function createPlot() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

    const spcDiv = d3.select(document.body)
                      .append('div')
                      .classed('spc-container', true)
                      .attr("hidden", true)
                      .node();

    const spcVisual = new Visual(makeConstructorArgs(spcDiv));

    var updateArgs = {
      dataViews: dummyData,
      viewport: { width: 640, height: 480 },
      type: 2,
      headless: true,
      frontend: true
    };

    spcVisual.update(updateArgs as any);
    spcVisual.svg
              .append("rect")
              .attr("width", "100%")
              .attr("height", "100%")
              .attr("fill", "white")
              .lower();

    var image = currentWorksheet.shapes.addImage(btoa(spcVisual.svg.node().outerHTML));
    image.name = "Image";

    await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
