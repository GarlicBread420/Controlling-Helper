export function taskpaneUtils()
{ }

/**
 * open features in taskpane
 * 
 * @param {string} pTitle title of the feature (eg. "autoControlling")
 */
taskpaneUtils.openFeatures = function (pTitle: string)
{
    switch (pTitle) {
        case ("autoControlling"):
            window.location.href = "autoControlling_feature.html";
            break;
        case ("pieChart"):
            window.location.href = "pieChart_feature.html";
            break;
    }
}