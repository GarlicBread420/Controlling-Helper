import { featureUtils } from "../../library/featureUtils";
Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("backButton").onclick = featureUtils.goBack;
  }
});