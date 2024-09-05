/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { taskpaneUtils } from "../library/taskpaneUtils";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("autoControlling_feature_button").onclick = () =>
      taskpaneUtils.openFeatures("autoControlling");
    document.getElementById("pieChart_feature_button").onclick = () =>
      taskpaneUtils.openFeatures("pieChart");
  }
});
